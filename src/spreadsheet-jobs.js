'use strict';
const SpreadsheetJobs = (() => {
  const MAX_FILE_BYTES = 5 * 1024 * 1024;
  const workerUrl = new URL('spreadsheet-worker.js', document.currentScript.src);
  function run(task, { signal, timeoutMs = 15000 } = {}) {
    return new Promise((resolve, reject) => {
      if (task.type === 'read' && task.file.size > MAX_FILE_BYTES) {
        reject(new Error('This file exceeds the 5 MB limit. Choose a smaller spreadsheet.'));
        return;
      }
      if (signal?.aborted) {
        reject(new DOMException('Cancelled', 'AbortError'));
        return;
      }
      let worker;
      try {
        worker = new Worker(workerUrl);
      } catch {
        reject(
          new Error(
            'Spreadsheet processing could not start. Open the app through its local server or hosted URL.',
          ),
        );
        return;
      }
      let timer,
        finished = false;
      const finish = (error, result) => {
        if (finished) return;
        finished = true;
        clearTimeout(timer);
        worker.terminate();
        signal?.removeEventListener('abort', abort);
        if (error) reject(error);
        else resolve(result);
      };
      const abort = () => finish(new DOMException('Cancelled', 'AbortError'));
      worker.onmessage = (event) => {
        if (event.data.ok) finish(null, event.data.result);
        else finish(new Error(event.data.error || 'Could not process this spreadsheet.'));
      };
      worker.onerror = (event) => {
        event.preventDefault();
        finish(new Error('Spreadsheet processing failed. Check the file and try again.'));
      };
      worker.onmessageerror = () =>
        finish(new Error('Could not read spreadsheet processing results.'));
      signal?.addEventListener('abort', abort, { once: true });
      timer = setTimeout(
        () =>
          finish(
            new Error(
              'Spreadsheet processing exceeded 15 seconds and was stopped. Try a smaller file.',
            ),
          ),
        timeoutMs,
      );
      try {
        worker.postMessage(task);
      } catch {
        finish(new Error('Could not start spreadsheet processing.'));
      }
    });
  }
  return { run, MAX_FILE_BYTES };
})();
