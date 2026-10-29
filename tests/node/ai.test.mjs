import test from 'node:test';
import assert from 'node:assert/strict';
import { handleRequest, validateCases, validateResolutions } from '../../server/ai-worker.js';

const accessCode = 'synthetic-test-access-code-123456';
const example = () => ({
  cell: 'C21',
  text: 'Take over Jaydon L from Luca at 13:30',
  reason: 'incoming instructor is missing or ambiguous',
  time: '13:30',
  instructors: ['adrian', 'luca'],
  appointments: [{ id: 'a0', student: 'Jaydon L', staff: 'luca', start: '13:00', end: '14:00' }],
  nearby: [
    { cell: 'B21', text: 'attention', color: 'rgb:FCE5CD', instructor: 'adrian', time: 790 },
  ],
});
const suggestion = () => ({
  cell: 'C21',
  confident: true,
  incomingInstructor: 'adrian',
  time: '13:30',
  studentIds: ['a0'],
});
const request = (body = { cases: [example()] }, headers = {}) =>
  new Request('https://schedule.example/api/handoffs', {
    method: 'POST',
    headers: {
      Origin: 'https://schedule.example',
      Authorization: `Bearer ${accessCode}`,
      'Content-Type': 'application/json',
      ...headers,
    },
    body: JSON.stringify(body),
  });
const environment = () => ({
  AI_ENABLED: 'true',
  OPENAI_API_KEY: 'fake-provider-key',
  OPENAI_MODEL: 'test-model',
  AI_ACCESS_TOKEN: accessCode,
  AI_RATE_LIMITER: { limit: async () => ({ success: true }) },
});
const never = () => {
  throw new Error('Unexpected provider request');
};

test('AI defaults off and requires every deployment credential', async () => {
  for (const field of [
    'AI_ENABLED',
    'OPENAI_API_KEY',
    'OPENAI_MODEL',
    'AI_ACCESS_TOKEN',
    'AI_RATE_LIMITER',
  ]) {
    const env = environment();
    delete env[field];
    assert.equal((await handleRequest(request(), env, never)).status, 503);
  }
});
test('AI route rejects foreign origins, invalid access codes, and excess requests', async () => {
  assert.equal(
    (
      await handleRequest(
        request(undefined, { Origin: 'https://foreign.example' }),
        environment(),
        never,
      )
    ).status,
    403,
  );
  assert.equal(
    (
      await handleRequest(
        request(undefined, { Authorization: 'Bearer wrong' }),
        environment(),
        never,
      )
    ).status,
    401,
  );
  const env = environment();
  env.AI_RATE_LIMITER.limit = async () => ({ success: false });
  assert.equal((await handleRequest(request(), env, never)).status, 429);
});
test('AI route enforces method, content type and streamed body size', async () => {
  assert.equal(
    (
      await handleRequest(
        new Request('https://schedule.example/api/handoffs'),
        environment(),
        never,
      )
    ).status,
    405,
  );
  assert.equal(
    (await handleRequest(request({}, { 'Content-Type': 'text/plain' }), environment(), never))
      .status,
    415,
  );
  assert.equal(
    (
      await handleRequest(
        request({ cases: [], extra: 'x'.repeat(128 * 1024) }),
        environment(),
        never,
      )
    ).status,
    400,
  );
  assert.equal((await handleRequest(request({ cases: [] }), environment(), never)).status, 400);
});
test('Only bounded schedule context is forwarded, with strict structured output and storage disabled', async () => {
  const item = example();
  item.binder = 'DO NOT SEND';
  const response = await handleRequest(
    request({ cases: [item] }),
    environment(),
    async (url, options) => {
      assert.equal(url, 'https://api.openai.com/v1/responses');
      const body = JSON.parse(options.body);
      assert.equal(body.store, false);
      assert.equal(body.text.format.strict, true);
      assert.equal(body.model, 'test-model');
      assert.ok(!body.input.includes('DO NOT SEND'));
      assert.ok(!body.input.includes(accessCode));
      return Response.json({
        status: 'completed',
        output: [
          {
            type: 'message',
            content: [
              { type: 'output_text', text: JSON.stringify({ resolutions: [suggestion()] }) },
            ],
          },
        ],
      });
    },
  );
  assert.equal(response.status, 200);
  assert.equal(response.headers.get('Cache-Control'), 'no-store');
  assert.deepEqual((await response.json()).resolutions, [
    {
      cell: 'C21',
      text: item.text,
      time: '13:30',
      incomingInstructor: 'adrian',
      studentIds: ['a0'],
    },
  ]);
});
test('AI cannot invent instructors, students, times, or duplicate case resolutions', () => {
  const cases = validateCases([example()]);
  for (const patch of [
    { incomingInstructor: 'invented' },
    { studentIds: ['a999'] },
    { studentIds: ['a0', 'a0'] },
    { time: '13:40' },
    { time: '25:00' },
  ])
    assert.throws(() => validateResolutions([{ ...suggestion(), ...patch }], cases));
  assert.throws(() => validateResolutions([suggestion(), suggestion()], cases));
  assert.deepEqual(validateResolutions([{ ...suggestion(), confident: false }], cases), []);
  const noTime = validateCases([{ ...example(), time: null }]);
  assert.throws(() => validateResolutions([{ ...suggestion(), time: '14:00' }], noTime));
});
test('Provider failure, refusal, incomplete and malformed replies fail safely without leaking errors', async () => {
  const outputs = [
    Response.json({ error: { message: 'secret-provider-detail' } }, { status: 500 }),
    Response.json({ status: 'incomplete' }),
    Response.json({
      status: 'completed',
      output: [{ type: 'message', content: [{ type: 'refusal' }] }],
    }),
    Response.json({
      status: 'completed',
      output: [{ type: 'message', content: [{ type: 'output_text', text: 'broken' }] }],
    }),
  ];
  for (const output of outputs) {
    const response = await handleRequest(request(), environment(), async () => output);
    assert.equal(response.status, 502);
    assert.ok(!(await response.text()).includes('secret-provider-detail'));
  }
});
test('Static assets remain served through their binding', async () => {
  const response = await handleRequest(
    new Request('https://schedule.example/'),
    { ASSETS: { fetch: () => new Response('static') } },
    never,
  );
  assert.equal(await response.text(), 'static');
});
