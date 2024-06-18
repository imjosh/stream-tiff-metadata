import fs from 'fs';
import path from 'path';
import { promisify } from 'util';

const readdir = promisify(fs.readdir);

const TIFF_TAGS = {
  256: 'ImageWidth',
  257: 'ImageLength',
  282: 'XResolution',
  283: 'YResolution'
};

const rootFolderPath = 'C:/path/to/your/tiffs';
main();


function readUInt(buffer, offset, length, isLittleEndian) {
  if (isLittleEndian) {
    return buffer.readUIntLE(offset, length);
  } else {
    return buffer.readUIntBE(offset, length);
  }
}

function readIFD(buffer, offset, isLittleEndian) {
  const entries = buffer.readUInt16LE(offset);
  const tags = {};

  for (let i = 0; i < entries; i++) {
    const entryOffset = offset + 2 + i * 12;
    const tag = readUInt(buffer, entryOffset, 2, isLittleEndian);
    const type = readUInt(buffer, entryOffset + 2, 2, isLittleEndian);
    const count = readUInt(buffer, entryOffset + 4, 4, isLittleEndian);
    const valueOffset = readUInt(buffer, entryOffset + 8, 4, isLittleEndian);

    if (TIFF_TAGS[tag]) {
      tags[TIFF_TAGS[tag]] = { type, count, valueOffset, entryOffset };
    }
  }

  return tags;
}

async function readTagValue(filePath, offset, length) {
  return new Promise((resolve, reject) => {
    const stream = fs.createReadStream(filePath, { start: offset, end: offset + length - 1 });
    let data = Buffer.alloc(0);

    stream.on('data', (chunk) => {
      data = Buffer.concat([data, chunk]);
    });

    stream.on('end', () => {
      resolve(data);
    });

    stream.on('error', (error) => {
      reject(error);
    });
  });
}

async function readResolution(filePath, offset, isLittleEndian) {
  const buffer = await readTagValue(filePath, offset, 8);
  const numerator = readUInt(buffer, 0, 4, isLittleEndian);
  const denominator = readUInt(buffer, 4, 4, isLittleEndian);
  return numerator / denominator;
}

async function readTiffMetadata(filePath) {
  return new Promise((resolve, reject) => {
    const stream = fs.createReadStream(filePath, { start: 0, end: 8 });

    let headerData = Buffer.alloc(0);

    stream.on('data', (chunk) => {
      headerData = Buffer.concat([headerData, chunk]);

      if (headerData.length >= 8) {
        const isLittleEndian = headerData.toString('utf8', 0, 2) === 'II';
        const magicNumber = readUInt(headerData, 2, 2, isLittleEndian);
        if (magicNumber !== 42) {
          reject(new Error('Not a valid TIFF file'));
          return;
        }

        const firstIFDOffset = readUInt(headerData, 4, 4, isLittleEndian);
        stream.destroy();

        // Read the number of entries in the IFD
        const entryCountStream = fs.createReadStream(filePath, { start: firstIFDOffset, end: firstIFDOffset + 1 });
        let entryCountData = Buffer.alloc(0);

        entryCountStream.on('data', (chunk) => {
          entryCountData = Buffer.concat([entryCountData, chunk]);
        });

        entryCountStream.on('end', () => {
          const entries = readUInt(entryCountData, 0, 2, isLittleEndian);
          const ifdSize = 2 + entries * 12 + 4; // 2 bytes for the entry count, 12 bytes per entry, 4 bytes for the next IFD offset

          const ifdStream = fs.createReadStream(filePath, { start: firstIFDOffset, end: firstIFDOffset + ifdSize - 1 });
          let ifdData = Buffer.alloc(0);

          ifdStream.on('data', (chunk) => {
            ifdData = Buffer.concat([ifdData, chunk]);

            if (ifdData.length >= ifdSize) {
              const tags = readIFD(ifdData, 0, isLittleEndian);

              // Determine if ImageWidth and ImageLength are stored directly or need to be read from offset
              const width = tags.ImageWidth.type === 3 ? tags.ImageWidth.valueOffset & 0xffff : readUInt(ifdData, tags.ImageWidth.entryOffset + 8, 4, isLittleEndian);
              const height = tags.ImageLength.type === 3 ? tags.ImageLength.valueOffset & 0xffff : readUInt(ifdData, tags.ImageLength.entryOffset + 8, 4, isLittleEndian);

              Promise.all([
                readResolution(filePath, tags.XResolution.valueOffset, isLittleEndian),
                readResolution(filePath, tags.YResolution.valueOffset, isLittleEndian)
              ]).then(([xResolution, yResolution]) => {
                resolve({ width, height, xResolution, yResolution });
                ifdStream.destroy();
              }).catch((error) => {
                reject(error);
                ifdStream.destroy();
              });
            }
          });

          ifdStream.on('error', (error) => {
            reject(error);
          });

          ifdStream.on('end', () => {
            if (ifdData.length < ifdSize) {
              reject(new Error('Unable to read TIFF metadata from the file.'));
            }
          });
        });

        entryCountStream.on('error', (error) => {
          reject(error);
        });

        entryCountStream.on('end', () => {
          if (entryCountData.length < 2) {
            reject(new Error('Unable to read the number of entries in the IFD.'));
          }
        });
      }
    });

    stream.on('error', (error) => {
      reject(error);
    });

    stream.on('end', () => {
      if (headerData.length < 8) {
        reject(new Error('Unable to read TIFF metadata from the file.'));
      }
    });
  });
}

async function processTiffFiles(directory) {
  try {
    const files = await readdir(directory);
    const tiffFiles = files.filter(file => file.endsWith('.tif'));

    for (const file of tiffFiles) {
      const filePath = path.join(directory, file);
      try {
        const metadata = await readTiffMetadata(filePath);
        // todo - write to a file
        console.log(`File: ${file}, Width: ${metadata.width}, Height: ${metadata.height}, DPI X: ${metadata.xResolution}, DPI Y: ${metadata.yResolution}`);
      } catch (error) {
        console.error(`Error reading metadata for file ${file}:`, error.message);
      }
    }
  } catch (error) {
    console.error(`Error processing TIFF files in ${directory}:`, error);
  }
}

async function main() {
  try {
    const subdirectories = await readdir(rootFolderPath, { withFileTypes: true });
    for (const dirent of subdirectories) {
      if (dirent.isDirectory()) {
        const subDirPath = path.join(rootFolderPath, dirent.name);
        await processTiffFiles(subDirPath);
      }
    }
  } catch (error) {
    console.error('Error processing directories:', error);
  }
}
