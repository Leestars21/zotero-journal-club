// Read-only PPTX preview through an existing @oai/artifact-tool runtime.
// Usage: node render_preview.mjs <artifact_tool.mjs> <input.pptx> <output-dir>
import fs from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';

const [runtimeModule, input, output] = process.argv.slice(2);
if (!runtimeModule || !input || !output) {
  console.error('Usage: node render_preview.mjs <artifact_tool.mjs> <input.pptx> <output-dir>');
  process.exit(2);
}
const { FileBlob, PresentationFile } = await import(pathToFileURL(path.resolve(runtimeModule)).href);
const deck = await PresentationFile.importPptx(await FileBlob.load(path.resolve(input)));
await fs.mkdir(output, { recursive: true });
for (let index = 0; index < deck.slides.items.length; index++) {
  const target = path.join(output, `slide-${String(index + 1).padStart(2, '0')}.png`);
  try {
    await fs.access(target);
    throw new Error(`Refusing to overwrite existing preview: ${target}`);
  } catch (error) {
    if (error.code !== 'ENOENT') throw error;
  }
}
for (let index = 0; index < deck.slides.items.length; index++) {
  const image = await deck.slides.items[index].export({ format: 'png', scale: 1.5 });
  await fs.writeFile(path.join(output, `slide-${String(index + 1).padStart(2, '0')}.png`), new Uint8Array(await image.arrayBuffer()));
}
console.log(`Rendered ${deck.slides.items.length} slides; inspect every page visually.`);
