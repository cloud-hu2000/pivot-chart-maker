import assert from 'node:assert/strict';
import { existsSync, readFileSync } from 'node:fs';
const required = [
  '/pivot-min-max-distinct',
  '/pivot-percent-change',
  '/pivot-two-row-dimensions',
  '/excel-to-csv',
  '/count-unique-values-csv',
];

function assertHub(label, html) {
  for (const href of required) {
    assert.match(html, new RegExp(`href="${href}"`), `${label} missing ${href}`);
  }
  const more = html.match(/<section class="cards" aria-label="More data tools">([\s\S]*?)<\/section>/);
  assert.ok(more, `${label} missing More data tools section`);
  const excel = more[1].indexOf('href="/excel-to-csv"');
  const min = more[1].indexOf('href="/pivot-min-max-distinct"');
  const percent = more[1].indexOf('href="/pivot-percent-change"');
  const two = more[1].indexOf('href="/pivot-two-row-dimensions"');
  assert.ok(excel >= 0 && excel < min && min < percent && percent < two, `${label} card order`);
}

assertHub('tools.html', readFileSync(new URL('../../tools.html', import.meta.url), 'utf8'));
const distUrl = new URL('../../dist/tools.html', import.meta.url);
if (existsSync(distUrl)) {
  assertHub('dist/tools.html', readFileSync(distUrl, 'utf8'));
  console.log('tools index linkage: source and dist hubs include the pivot cards');
} else {
  console.log('tools index linkage: source hub includes the pivot cards; dist/tools.html is not built yet');
}
