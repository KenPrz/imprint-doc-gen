// Independent oracle: npm install --prefix /tmp/imprint-carbone-research-20261018 --ignore-scripts carbone@3.8.2
// node tests/fixtures/migration/carbone-probe.cjs /tmp/imprint-carbone-research-20261018/node_modules/carbone
const carbone = require(process.argv[2]);
const cases = [
  ['basic', 'Hello {d.firstname} {d.lastname} !', { firstname: 'John', lastname: 'Doe' }],
  ['nested', '{d.person.name}', { person: { name: 'Example' } }],
  ['missing', '{d.name}', {}], ['null', '{d.name}', { name: null }],
  ['false', '{d.name}', { name: false }], ['zero', '{d.name}', { name: 0 }],
  ['upper', '{d.name:upperCase}', { name: 'Straße café σ' }],
  ['lower', '{d.name:lowerCase}', { name: 'İSTANBUL CAFÉ Σ' }],
  ['chain', '{d.name:upperCase:lowerCase}', { name: 'AbC' }],
  ['upper-null', '{d.name:upperCase}', { name: null }],
  ['upper-number', '{d.name:upperCase}', { name: 123 }],
  ['lower-false', '{d.name:lowerCase}', { name: false }],
];
(async () => {
 const outputs = [];
 for (const [name, source, data] of cases) {
  const output = await new Promise((resolve, reject) => carbone.renderXML(source, data, { extension: 'txt' }, (error, result) => error ? reject(error) : resolve(result)));
  outputs.push({ name, source, data, output });
 }
 process.stdout.write(JSON.stringify({ oracle: 'carbone@3.8.2 renderXML extension=txt', cases: outputs }, null, 2) + '\n');
})();
