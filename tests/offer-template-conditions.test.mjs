import assert from 'node:assert/strict';
import { test } from 'node:test';
import AdmZip from 'adm-zip';
import conditions from '../server/offer-template-conditions.js';

const template = new AdmZip('server/templates/tilbud/tilbud-stromskinner-template.docx');
const documentXml = template.readAsText('word/document.xml');

test('montasje conditions are removed with their leading page break', ()=>{
  const output = conditions.removeMontasjeConditionsPage(documentXml);
  assert.equal(output.includes('Forbehold ved montasje av strømskinner fra'), false);
  assert.equal(output.includes('Forutsetninger for pris'), false);
  assert.equal(output.includes('Betingelser'), true);
  assert.equal(output.includes('GENERELLE BETINGELER'), true);
  assert.equal((output.match(/<w:br w:type="page"\/>/g) || []).length,
    (documentXml.match(/<w:br w:type="page"\/>/g) || []).length - 1);
});

test('templates without a montasje conditions heading are unchanged', ()=>{
  const output = conditions.removeMontasjeConditionsPage('<w:document><w:body><w:p/></w:body></w:document>');
  assert.equal(output, '<w:document><w:body><w:p/></w:body></w:document>');
});
