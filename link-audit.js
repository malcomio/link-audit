const fs = require('fs');
const axios = require('axios');
const cheerio = require('cheerio');
const ExcelJS = require('exceljs');
require('dotenv').config();

const authUser = process.env.BASIC_AUTH_USER || '';
const authPass = process.env.BASIC_AUTH_PASSWORD || '';
const domain = `${process.env.TEST_DOMAIN}`;
const baseUrl = `https://${domain}`;

const MAX_PAGES = Number(process.env.CRAWL_MAX_PAGES || 500);
const REQUEST_DELAY_MS = Number(process.env.CRAWL_DELAY_MS || 250);

const visited = new Set();
const queue = [{ path: '/', sourcePage: '', linkText: '' }];
const queuedPaths = new Set(['/']);
const findings = [];
const errors = [];
const invalidLinks = [];
const redirectedLinks = [];
const analysedFinalPaths = new Set();

const sleep = ms => new Promise(resolve => setTimeout(resolve, ms));

function normaliseUrl(href, pageUrl) {
  try {
    const resolved = new URL(href, pageUrl);
    resolved.hash = '';
    return resolved;
  } catch {
    return null;
  }
}

function isSameSite(url) {
  if (typeof url === 'string') {
    if (url.startsWith('/')) {
      return true;
    }
  } 

  if (url.hostname === domain) {
    return true;
  }
  
  if (url.protocol === 'javascript:') {
    return true;
  }

  return false;
}

function isCrawlablePath(pathname) {
  // Skip obvious non-HTML resources.
  return !/\.(pdf|docx?|xlsx?|pptx?|csv|zip|jpg|jpeg|png|gif|svg|webp|ico|css|js|xml|json)$/i.test(pathname);
}

function getRobotsDirectives($) {
  let content = '';
  $('meta').each((_, el) => {
    if (($(el).attr('name') || '').toLowerCase() === 'robots') {
      content = $(el).attr('content') || '';
    }
  });
  return content.toLowerCase().split(',').map(s => s.trim()).filter(Boolean);
}

function isLikelyEmailAddress(value) {
  const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (emailPattern.test(value) && !value.startsWith('https://')) {    
    return true;
  }
  return false;
}

function extractDomain(href, pageUrl) {
  const value = (href || '').trim();
  if (!value) {
    return '';
  }

  if (isLikelyEmailAddress(value)) {
    return value.split('@')[1] || '';
  }

  try {
    return new URL(value, pageUrl).hostname;
  } catch {
    return '';
  }
}

function hasValidHttpHostname(hostname) {
  if (!hostname) {
    return false;
  }

  if (hostname === 'localhost') {
    return true;
  }

  if (/^\d{1,3}(\.\d{1,3}){3}$/.test(hostname)) {
    return true;
  }

  return hostname.includes('.');
}

function getInvalidLinkReason(rawHref) {
  const href = (rawHref || '').trim();

  if (!href || href.startsWith('#')) {
    return null;
  }

  if (href.toLowerCase().startsWith('mailto:')) {
    return null;
  }

  if (isLikelyEmailAddress(href)) {
    return 'Email address is not a mailto: link';
  }

  if (/^void\s*\(/i.test(href)) {
    return 'Invalid pseudo-link value';
  }

  if (/^https?:\/(?!\/)/i.test(href)) {
    return 'Invalid protocol format (expected http:// or https://)';
  }

  if (/^https?:\/\//i.test(href)) {
    try {
      const parsed = new URL(href);
      if (!hasValidHttpHostname(parsed.hostname)) {
        return 'Invalid domain in absolute URL';
      }
    } catch {
      return 'Invalid absolute URL';
    }
  }

  return null;
}

function shouldSkipForCrawling(href) {
  const value = (href || '').trim().toLowerCase();
  if (!value) {
    return true;
  }

  return value.startsWith('mailto:')
    || value.startsWith('tel:')
    || value.startsWith('javascript:')
    || /^void\s*\(/i.test(value);
}

async function fetchPage(path) {
  const url = `${baseUrl}${path}`;
  return axios.get(url, {
    auth: { username: authUser, password: authPass },
    timeout: 15000,
    validateStatus: () => true,
    headers: { 'User-Agent': 'dxp-testing target-blank audit' },
  });
}

function getFinalUrlFromResponse(response) {
  return response?.request?.res?.responseUrl || '';
}

function toPathWithSearch(inputUrl) {
  try {
    const parsed = new URL(inputUrl);
    return `${parsed.pathname}${parsed.search}`;
  } catch {
    return '';
  }
}

function toAbsoluteUrl(path) {
  return `${baseUrl}${path}`;
}

function sanitiseForExcel(value) {
  if (value === null || value === undefined) {
    return '';
  }

  return String(value)
    // Remove control characters not allowed in XML 1.0 (used by XLSX parts).
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '');
}

function buildSheetWithTable(workbook, sheetName, tableName, columns, rowObjects) {
  const sheet = workbook.addWorksheet(sheetName);
  const rows = rowObjects.map(row => columns.map(column => sanitiseForExcel(row[column])));

  sheet.views = [{ state: 'frozen', ySplit: 1 }];

  sheet.addTable({
    name: tableName,
    ref: 'A1',
    headerRow: true,
    totalsRow: false,
    style: {
      theme: 'TableStyleMedium2',
      showRowStripes: true,
    },
    columns: columns.map(column => ({ name: column })),
    rows,
  });

  // Auto-size columns for readability in exported reports.
  columns.forEach((column, index) => {
    const contentWidth = Math.max(
      column.length,
      ...rows.map(r => String(r[index] || '').length),
    );
    sheet.getColumn(index + 1).width = Math.min(80, contentWidth + 2);
  });

  return sheet;
}

async function crawl() {
  while (queue.length && visited.size < MAX_PAGES) {
    const current = queue.shift();
    const { path, sourcePage, linkText } = current;
    queuedPaths.delete(path);

    if (visited.has(path)) continue;
    visited.add(path);

    let response;
    try {
      response = await fetchPage(path);
    } catch (e) {
      errors.push({ path, error: e.message, sourcePage, linkText });
      continue;
    }

    if (response.status >= 400) {
      errors.push({ path, error: `HTTP ${response.status}`, sourcePage, linkText });
      continue;
    }

    const requestedUrl = toAbsoluteUrl(path);
    const finalUrl = getFinalUrlFromResponse(response) || requestedUrl;
    const finalPath = toPathWithSearch(finalUrl) || path;
    const canonicalPageUrl = toAbsoluteUrl(finalPath);

    if (finalPath !== path) {
      redirectedLinks.push({
        requestedPath: path,
        redirectedPath: finalPath,
        requestedUrl,
        redirectedUrl: finalUrl,
        sourcePage: sourcePage || '(seed)',
        linkText: linkText || '(n/a)',
      });
    }

    // Avoid duplicate findings when multiple URLs resolve to the same final page.
    if (analysedFinalPaths.has(finalPath)) {
      console.log(`Crawled (${visited.size}/${MAX_PAGES}): ${path} -> ${finalPath} (already analysed)`);
      await sleep(REQUEST_DELAY_MS);
      continue;
    }
    analysedFinalPaths.add(finalPath);

    const contentType = response.headers['content-type'] || '';
    if (!contentType.includes('text/html')) continue;

    const pageUrl = canonicalPageUrl;
    const $ = cheerio.load(response.data);

    const robotsDirectives = getRobotsDirectives($);
    const pageNoFollow = robotsDirectives.includes('nofollow');

    $('a[target="_blank"]').each((_, el) => {
      const rel = $(el).attr('rel') || '';
      if (!rel.includes('noopener') && !isSameSite($(el).attr('href'))) {
        findings.push({
          page: pageUrl,
          href: $(el).attr('href') || '',
          text: $(el).text().trim().slice(0, 80),
          rel: rel || '(none)',
        });
      }
    });

    if (!pageNoFollow) {
      $('a[href]').each((_, el) => {
        const rawHref = ($(el).attr('href') || '').trim();
        const linkTextValue = $(el).text().trim().slice(0, 120);

        const invalidReason = getInvalidLinkReason(rawHref);
        if (invalidReason) {
          invalidLinks.push({
            page: pageUrl,
            href: rawHref,
            text: linkTextValue,
            reason: invalidReason,
          });
        }

        const linkRel = ($(el).attr('rel') || '').toLowerCase();
        if (linkRel.includes('nofollow')) {
          return;
        }

        if (shouldSkipForCrawling(rawHref)) {
          return;
        }

        const resolved = normaliseUrl(rawHref, pageUrl);
        if (!resolved) {
          return;
        }

        if (!isSameSite(resolved)) {
          return;
        }

        if (!isCrawlablePath(resolved.pathname)) {
          return;
        }

        const nextPath = resolved.pathname + resolved.search;
        if (!visited.has(nextPath) && !queuedPaths.has(nextPath)) {
          queue.push({
            path: nextPath,
            sourcePage: canonicalPageUrl,
            linkText: linkTextValue,
          });
          queuedPaths.add(nextPath);
        }
      });
    }

    console.log(`Crawled (${visited.size}/${MAX_PAGES}): ${path}`);
    await sleep(REQUEST_DELAY_MS);
  }
}

(async () => {
  console.log(`Starting crawl of ${baseUrl}`);
  await crawl();

  console.log(`\nPages crawled: ${visited.size}`);
  console.log(`Pages with errors: ${errors.length}`);
  console.log(`Redirected links found: ${redirectedLinks.length}`);
  console.log(`Unsafe target="_blank" links found: ${findings.length}`);
  console.log(`Invalid links found: ${invalidLinks.length}\n`);

  const reportDir = 'reports';
  if (!fs.existsSync(reportDir)) fs.mkdirSync(reportDir, { recursive: true });

  const findingsRows = findings.map(f => ({
    Page: f.page,
    Href: f.href,
    Domain: extractDomain(f.href, f.page),
    Rel: f.rel,
    Text: f.text,
  }));

  const failedPageRows = errors.map(e => ({
    Path: e.path,
    Error: e.error,
    'Found On Page': e.sourcePage || '(seed)',
    'Link Text': e.linkText || '(n/a)',
  }));

  const invalidLinkRows = invalidLinks.map(link => ({
    Page: link.page,
    Href: link.href,
    Text: link.text,
    Issue: link.reason,
  }));

  const redirectedLinkRows = redirectedLinks.map(link => ({
    'Requested Path': link.requestedPath,
    'Redirected Path': link.redirectedPath,
    'Requested URL': link.requestedUrl,
    'Redirected URL': link.redirectedUrl,
    'Found On Page': link.sourcePage,
    'Link Text': link.linkText,
  }));

  const workbook = new ExcelJS.Workbook();

  buildSheetWithTable(
    workbook,
    'Unsafe target_blank links',
    'UnsafeTargetBlankLinks',
    ['Page', 'Href', 'Domain', 'Rel', 'Text'],
    findingsRows,
  );

  buildSheetWithTable(
    workbook,
    'Failed page loads',
    'FailedPageLoads',
    ['Path', 'Error', 'Found On Page', 'Link Text'],
    failedPageRows,
  );

  buildSheetWithTable(
    workbook,
    'Invalid links',
    'InvalidLinks',
    ['Page', 'Href', 'Domain', 'Text', 'Issue'],
    invalidLinkRows,
  );

  buildSheetWithTable(
    workbook,
    'Redirected links',
    'RedirectedLinks',
    ['Requested Path', 'Redirected Path', 'Requested URL', 'Redirected URL', 'Found On Page', 'Link Text'],
    redirectedLinkRows,
  );

  // Get today's date and time.
  const today = new Date();
  const dateStr = today.toISOString().split('.')[0].replaceAll(':','-'); // e.g., "2025-05-15T03:56:04"

  const reportPath = `${reportDir}/link-audit-${domain}-${dateStr}.xlsx`;
  await workbook.xlsx.writeFile(reportPath);
  console.log(`\nXLSX report written to ${reportPath}`);

  process.exit(findings.length || invalidLinks.length ? 1 : 0);
})();
