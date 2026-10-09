/**
 * images.js — the image store.
 *
 * An image lives in two places: a short reference token inside the source
 * text, which is what gives it a *position* in the document, and its bytes
 * here, keyed by that reference. Keeping the bytes out of the text is what
 * lets the source pane stay readable and lets the parser treat an image as an
 * ordinary block that content flows around.
 *
 * Imported images are downscaled before they are stored. A phone photo is
 * several megapixels and nothing on an A4 page can use more than a fraction
 * of that; storing the original would blow the localStorage quota and make
 * every fit measurement slower for no visible gain.
 */

const STORE_KEY = 'onepage.images.v1';

/** Longest edge kept for a stored image. 1600px is well past 300dpi on A4. */
const MAX_EDGE = 1600;
const JPEG_QUALITY = 0.86;

/** @typedef {{src:string, width:number, height:number, name:string}} StoredImage */

/** @type {Map<string, StoredImage>} */
const store = new Map();

let counter = 0;

export function newId() {
  counter += 1;
  return `${Date.now().toString(36)}${counter.toString(36)}`;
}

/** @returns {StoredImage|null} */
export const get = (id) => store.get(id) || null;

export const has = (id) => store.has(id);

export const put = (id, image) => {
  store.set(id, image);
};

export const remove = (id) => store.delete(id);

export const clear = () => store.clear();

/** Drop anything the document no longer references. */
export function keepOnly(ids) {
  const wanted = new Set(ids);
  for (const id of [...store.keys()]) if (!wanted.has(id)) store.delete(id);
}

function readAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(new Error('Could not read the image file'));
    reader.readAsDataURL(file);
  });
}

function load(src) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = () => reject(new Error('Could not decode the image'));
    img.src = src;
  });
}

/**
 * Downscale to MAX_EDGE. A PNG with transparency stays a PNG; everything else
 * becomes a JPEG, because a photograph re-encoded as PNG is many times larger
 * than the original for no benefit.
 */
function downscale(img, mime) {
  const scale = Math.min(1, MAX_EDGE / Math.max(img.naturalWidth, img.naturalHeight));
  const width = Math.max(1, Math.round(img.naturalWidth * scale));
  const height = Math.max(1, Math.round(img.naturalHeight * scale));

  const canvas = document.createElement('canvas');
  canvas.width = width;
  canvas.height = height;
  const ctx = canvas.getContext('2d');
  ctx.imageSmoothingQuality = 'high';

  const keepAlpha = mime === 'image/png' || mime === 'image/webp' || mime === 'image/gif';
  if (!keepAlpha) {
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, width, height);
  }
  ctx.drawImage(img, 0, 0, width, height);

  const src = keepAlpha ? canvas.toDataURL('image/png') : canvas.toDataURL('image/jpeg', JPEG_QUALITY);
  return { src, width, height };
}

/**
 * Import a File (or Blob) and register it.
 *
 * @param {File|Blob} file
 * @returns {Promise<{id:string, image:StoredImage}>}
 */
export async function importFile(file) {
  if (!file || !String(file.type || '').startsWith('image/')) {
    throw new Error('That file is not an image');
  }
  // SVG has no raster dimensions to downscale and re-encoding it through a
  // canvas would throw away its sharpness, so it is stored as-is.
  const original = await readAsDataUrl(file);
  const img = await load(original);

  const isSvg = file.type === 'image/svg+xml';
  const sized = isSvg
    ? { src: original, width: img.naturalWidth || 400, height: img.naturalHeight || 300 }
    : downscale(img, file.type);

  // A re-encode that came out larger than the original is not worth keeping.
  const src = !isSvg && sized.src.length > original.length ? original : sized.src;

  const id = newId();
  const image = { src, width: sized.width, height: sized.height, name: file.name || 'image' };
  store.set(id, image);
  return { id, image };
}

/* --- Persistence -------------------------------------------------------- */

export function save() {
  const obj = {};
  for (const [id, image] of store) obj[id] = image;
  try {
    localStorage.setItem(STORE_KEY, JSON.stringify(obj));
    return true;
  } catch {
    // Over quota. The document still works for this session; it is the
    // images, not the text, that will be missing after a reload.
    try {
      localStorage.removeItem(STORE_KEY);
    } catch {
      /* storage unavailable entirely */
    }
    return false;
  }
}

export function restore() {
  try {
    const raw = localStorage.getItem(STORE_KEY);
    if (!raw) return;
    const obj = JSON.parse(raw);
    for (const [id, image] of Object.entries(obj || {})) {
      if (image && typeof image.src === 'string') {
        store.set(id, {
          src: image.src,
          width: Number(image.width) || 400,
          height: Number(image.height) || 300,
          name: String(image.name || 'image'),
        });
      }
    }
  } catch {
    /* corrupt or unavailable storage */
  }
}
