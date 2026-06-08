import sharp from "sharp";
import { resolve, dirname } from "path";
import { fileURLToPath } from "url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const root = resolve(__dirname, "..");
const sourceImage = resolve(root, "assets/icon-source.png");

const sizes = [16, 32, 64, 80, 128];
for (const size of sizes) {
  await sharp(sourceImage)
    .resize(size, size, { fit: "contain", background: { r: 0, g: 0, b: 0, alpha: 0 } })
    .png()
    .toFile(resolve(root, `assets/icon-${size}.png`));
  console.log(`generated icon-${size}.png`);
}
console.log("done");
