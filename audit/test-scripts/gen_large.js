const fs = require('fs');
const path = require('path');

function randWords(n) {
  const words = ['whereas','the','parties','agree','that','liability','shall','not','exceed','fees','paid','in','the','preceding','twelve','months','except','for','breaches','of','confidentiality','indemnification','obligations','gross','negligence','or','willful','misconduct','each','party','warrants','and','represents','data','protection','governing','law','termination','notice','period','service','levels','change','management','force','majeure','intellectual','property','rights','subcontractor','audit'];
  let out = [];
  for (let i=0;i<n;i++) out.push(words[Math.floor(Math.random()*words.length)]);
  return out.join(' ');
}

const categories = ['Limitation of Liability','Indemnity','Confidentiality','Data Protection & Privacy','Termination','Payment Terms'];
const statuses = ['Preferred','Fallback','Retired','Unclassified'];
const positions = ['provider','balanced','customer','none'];
const types = ['Clause','Comment','Clause + Comment'];

const N = parseInt(process.argv[2] || '2000', 10);
const withRichPct = parseFloat(process.argv[3] || '0.3');

const entries = [];
for (let i = 0; i < N; i++) {
  const clauseText = randWords(120 + Math.floor(Math.random()*200));
  const hasRich = Math.random() < withRichPct;
  let wordOpenXml = '';
  if (hasRich) {
    // simulate a moderately-sized WordOpenXML payload (real ones can be 5-50x plain text due to style/rsid bloat)
    wordOpenXml = '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">' + randWords(1500) + '</pkg:package>';
  }
  entries.push({
    id: 'seed-' + i + '-' + Math.random().toString(16).slice(2),
    title: 'Clause ' + i + ' — ' + categories[i % categories.length],
    type: types[i % types.length],
    category: categories[i % categories.length],
    status: statuses[i % statuses.length],
    position: positions[i % positions.length],
    favorite: i % 11 === 0,
    clauseText,
    commentText: i % 3 === 0 ? randWords(40) : '',
    contextText: i % 5 === 0 ? randWords(20) : '',
    wordOpenXml,
    richSourceText: hasRich ? clauseText : '',
    tags: ['tag' + (i % 20), 'tag' + (i % 7)],
    source: 'Synthetic MSA ' + (i % 50),
    needsReview: i % 13 === 0,
    createdAt: new Date(Date.now() - i * 60000).toISOString(),
    updatedAt: new Date(Date.now() - i * 30000).toISOString(),
    entryRevision: 1 + (i % 4)
  });
}

const payload = {
  app: 'ClauseCommentsLibrary', version: 6, libraryId: 'lib-synthetic-' + N,
  libraryName: 'Synthetic Scale Test Library', revision: 5, savedAt: new Date().toISOString(),
  entries, inboxSeenIds: []
};

const outPath = path.join(__dirname, 'out', `synthetic_${N}.json`);
fs.writeFileSync(outPath, JSON.stringify(payload));
console.log('Wrote', outPath, 'size bytes:', fs.statSync(outPath).size, 'entries:', N, 'richPct:', withRichPct);
