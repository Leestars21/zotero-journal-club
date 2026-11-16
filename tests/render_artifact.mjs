// Read-only preview of a PPTX using the installed artifact runtime.
import fs from 'node:fs/promises';
import { pathToFileURL } from 'node:url';

const [runtimeModule, input, output] = process.argv.slice(2);
const { FileBlob, PresentationFile } = await import(pathToFileURL(runtimeModule).href);
const deck = await PresentationFile.importPptx(await FileBlob.load(input));
await fs.mkdir(output, { recursive: true });
for (let i = 0; i < deck.slides.items.length; i++) {
  const blob = await deck.slides.items[i].export({ format: 'png', scale: 1.5 });
  await fs.writeFile(`${output}/slide-${i + 1}.png`, new Uint8Array(await blob.arrayBuffer()));
}
console.log(`Rendered ${deck.slides.items.length} slides`);
