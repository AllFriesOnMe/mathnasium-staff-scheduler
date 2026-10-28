import { copyFileSync, existsSync, lstatSync, mkdirSync, rmSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

export const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
export const publicFiles = Object.freeze([
  'index.html',
  'src/styles.css',
  'src/scheduler.js',
  'src/spreadsheet.js',
  'src/spreadsheet-jobs.js',
  'src/spreadsheet-worker.js',
  '_headers',
  'vendor/xlsx.full.min.js',
  'vendor/LICENSE-sheetjs.txt',
]);
export function build() {
  const output = resolve(root, 'dist');
  if (output !== join(root, 'dist')) throw new Error('Unexpected build output path.');
  // Never follow symlinks/junctions when copying or replacing deployment files.
  for (const file of publicFiles) {
    const source = join(root, file);
    for (const path of [dirname(source), source]) {
      if (lstatSync(path).isSymbolicLink())
        throw new Error(`Refusing linked deployment source: ${file}`);
    }
    if (!lstatSync(source).isFile()) throw new Error(`Not a regular deployment file: ${file}`);
  }
  if (existsSync(output)) {
    if (lstatSync(output).isSymbolicLink() || !lstatSync(output).isDirectory())
      throw new Error('dist must be a regular directory.');
    rmSync(output, { recursive: true });
  }
  mkdirSync(output);
  for (const file of publicFiles) {
    const target = join(output, file);
    mkdirSync(dirname(target), { recursive: true });
    copyFileSync(join(root, file), target);
  }
  return output;
}
if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  console.log(`Built ${publicFiles.length} approved files into ${build()}`);
}
