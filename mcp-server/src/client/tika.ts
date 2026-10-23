// SPDX-License-Identifier: MIT

import { logger } from '../logger.js';

/**
 * Apache Tika client for server-side document text extraction.
 *
 * read_file alone can only return text files: WebDAV fetches bytes and they are
 * decoded as UTF-8, so a .docx (a zip) comes back as unusable mojibake. Tika
 * turns office, OpenDocument and PDF files into plain text, and OCRs scans when
 * the Tika image bundles Tesseract (the `-full` tags do).
 *
 * Optional. With TIKA_URL unset, nothing here is used and read_file keeps its
 * previous behaviour for text files.
 */

const TIKA_URL = process.env.TIKA_URL?.replace(/\/+$/, '');
const TIKA_TIMEOUT_MS = parseInt(process.env.TIKA_TIMEOUT_MS || '', 10) || 120_000;
const TIKA_OCR_STRATEGY = process.env.TIKA_OCR_STRATEGY || 'auto';
// Extraction loads the whole file into Tika's heap; very large files stall it.
const TIKA_MAX_BYTES = parseInt(process.env.TIKA_MAX_BYTES || '', 10) || 50 * 1024 * 1024;

export function isTikaConfigured(): boolean {
  return Boolean(TIKA_URL);
}

export function getTikaMaxBytes(): number {
  return TIKA_MAX_BYTES;
}

/** Extensions read as plain text directly, without round-tripping through Tika. */
const TEXT_EXTENSIONS = new Set([
  'txt',
  'md',
  'markdown',
  'csv',
  'tsv',
  'log',
  'json',
  'yaml',
  'yml',
  'toml',
  'ini',
  'cfg',
  'conf',
  'env',
  'xml',
  'html',
  'htm',
  'css',
  'js',
  'mjs',
  'cjs',
  'ts',
  'tsx',
  'jsx',
  'py',
  'rb',
  'php',
  'go',
  'rs',
  'java',
  'c',
  'h',
  'cpp',
  'hpp',
  'cs',
  'sh',
  'bash',
  'zsh',
  'ps1',
  'sql',
  'svg',
  'rst',
  'tex',
  'org',
]);

/** Extensions Tika can turn into text. Anything else is refused rather than mangled. */
const EXTRACTABLE_MIME: Record<string, string> = {
  pdf: 'application/pdf',
  doc: 'application/msword',
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  xls: 'application/vnd.ms-excel',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  ppt: 'application/vnd.ms-powerpoint',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  odt: 'application/vnd.oasis.opendocument.text',
  ods: 'application/vnd.oasis.opendocument.spreadsheet',
  odp: 'application/vnd.oasis.opendocument.presentation',
  rtf: 'application/rtf',
  epub: 'application/epub+zip',
  msg: 'application/vnd.ms-outlook',
  eml: 'message/rfc822',
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  tif: 'image/tiff',
  tiff: 'image/tiff',
  gif: 'image/gif',
  bmp: 'image/bmp',
  webp: 'image/webp',
};

function extensionOf(path: string): string {
  const base = path.split('/').pop() ?? '';
  const dot = base.lastIndexOf('.');
  return dot === -1 ? '' : base.slice(dot + 1).toLowerCase();
}

export function isTextLike(path: string): boolean {
  const ext = extensionOf(path);
  return ext === '' || TEXT_EXTENSIONS.has(ext);
}

export function isExtractable(path: string): boolean {
  return extensionOf(path) in EXTRACTABLE_MIME;
}

export function mimeTypeFor(path: string): string {
  return EXTRACTABLE_MIME[extensionOf(path)] ?? 'application/octet-stream';
}

/**
 * Send bytes to Tika and get plain text back.
 *
 * Content-Type must be set explicitly: without it Tika sees
 * application/x-www-form-urlencoded, selects EmptyParser and returns an empty
 * body with HTTP 200 — a silent no-op rather than an error.
 */
export async function extractText(data: Uint8Array, path: string): Promise<string> {
  if (!TIKA_URL) {
    throw new Error('TIKA_URL is not configured; cannot extract text from this file type.');
  }

  const started = Date.now();
  const response = await fetch(`${TIKA_URL}/tika`, {
    method: 'PUT',
    headers: {
      'Content-Type': mimeTypeFor(path),
      Accept: 'text/plain',
      'X-Tika-PDFocrStrategy': TIKA_OCR_STRATEGY,
    },
    body: data as BodyInit,
    signal: AbortSignal.timeout(TIKA_TIMEOUT_MS),
  });

  if (!response.ok) {
    throw new Error(`Tika extraction failed: ${response.status} ${response.statusText}`);
  }

  const text = await response.text();
  logger.debug(
    { path, bytes: data.byteLength, chars: text.length, ms: Date.now() - started },
    '[tika] extracted'
  );
  return text;
}
