function getXmlElementRanges(xml, tagName) {
  const source = String(xml || '');
  const escapedTag = tagName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const tagPattern = new RegExp(`<\\/?${escapedTag}\\b[^>]*>`, 'g');
  const ranges = [];
  const stack = [];
  let match;
  while ((match = tagPattern.exec(source)) !== null) {
    const token = match[0];
    if (token.startsWith('</')) {
      const start = stack.pop();
      if (start !== undefined && stack.length === 0) {
        ranges.push({ start, end: tagPattern.lastIndex });
      }
    } else if (!token.endsWith('/>')) {
      stack.push(match.index);
    }
  }
  return ranges;
}

function removeMontasjeConditionsPage(xml) {
  const source = String(xml || '');
  const body = getXmlElementRanges(source, 'w:body')[0];
  if (!body) return source;

  const tables = getXmlElementRanges(source, 'w:tbl');
  const paragraphs = getXmlElementRanges(source, 'w:p')
    .filter(range=>range.start > body.start && range.end < body.end)
    .filter(range=>!tables.some(table=>range.start > table.start && range.end < table.end));
  const text = range=>[...source.slice(range.start, range.end).matchAll(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g)]
    .map(match=>match[1]).join('').replace(/\s+/g, ' ').trim();
  const headingIndex = paragraphs.findIndex(range=>
    /^Forbehold ved montasje av strømskinner fra\s*Busbar$/i.test(text(range))
  );
  if (headingIndex < 0) return source;

  const pageBreak = range=>/<w:br\b[^>]*\bw:type=["']page["'][^>]*\/?\s*>/.test(
    source.slice(range.start, range.end)
  );
  const start = paragraphs[headingIndex - 1];
  const nextHeadingIndex = paragraphs.findIndex((range, index)=>
    index > headingIndex && text(range) === 'Betingelser'
  );
  const end = paragraphs[nextHeadingIndex - 1];
  if (!start || !pageBreak(start) || !end || !pageBreak(end)) {
    throw new Error('Kunne ikke avgrense montasjeforbeholdet i tilbudsmalen');
  }
  return source.slice(0, start.start) + source.slice(end.start);
}

module.exports = { getXmlElementRanges, removeMontasjeConditionsPage };
