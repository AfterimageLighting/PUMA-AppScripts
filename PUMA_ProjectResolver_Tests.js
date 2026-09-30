/**
 * PUMA resolver/importer unit tests.
 *
 * These tests are read-only and do not access or modify spreadsheet data.
 * Run pumaRunResolverUnitTests() from Apps Script after pushing this branch to
 * a TEST Apps Script project, never as the first validation step in production.
 */

function pumaRunResolverUnitTests() {
  var results = [];

  function assertEqual(name, actual, expected) {
    var ok = actual === expected;
    results.push({name: name, ok: ok, actual: actual, expected: expected});
    if (!ok) throw new Error(name + ': expected "' + expected + '" but got "' + actual + '"');
  }

  function assertTrue(name, value) {
    var ok = !!value;
    results.push({name: name, ok: ok, actual: value, expected: true});
    if (!ok) throw new Error(name + ': expected true');
  }

  assertEqual(
    'QBO parent prefix + Residence normalization',
    pumaNormalizeProjectKey_('Arcadia:Kinney Residence'),
    'kinney'
  );

  assertEqual(
    'Residence optional for comparison',
    pumaNormalizeProjectKey_('Harrison Smith'),
    pumaNormalizeProjectKey_('Harrison Smith Residence')
  );

  assertEqual(
    'Dock Pickup is context, not identity',
    pumaNormalizeProjectKey_('Campbell Residence (Dock Pickup)'),
    'campbell'
  );

  assertEqual(
    'Housing Sample is context, not identity',
    pumaNormalizeProjectKey_('Stony Shore Residence (Housing Sample)'),
    'stony shore'
  );

  assertTrue(
    'Scheduled is protected from PO rollback',
    !canPoImporterSetOrdered_('Scheduled')
  );

  assertTrue(
    'Received is protected from PO rollback',
    !canPoImporterSetOrdered_('Received')
  );

  assertTrue(
    'Delivered is protected from PO rollback',
    !canPoImporterSetOrdered_('Delivered')
  );

  assertTrue(
    'Omitted is protected from PO rollback',
    !canPoImporterSetOrdered_('Omitted')
  );

  assertTrue(
    'Approved may advance to Ordered',
    canPoImporterSetOrdered_('Approved')
  );

  assertTrue(
    'Freight is a non-product line',
    isNonProductPoLine_({
      itemName: 'Freight and Shipping',
      itemType: '',
      description: ''
    })
  );

  assertTrue(
    'Physical fixture is not a non-product line',
    !isNonProductPoLine_({
      itemName: 'E2SLB-OW',
      itemType: 'C1-2A-TRIM',
      description: '2 inch Adjustable LED Flangeless Square Trim'
    })
  );

  var trackerInfo = {
    rows: [
      {
        sheetRow: 10,
        type: 'C1-2A-TRIM',
        partNumber: 'E2SLB-OW',
        description: '2 inch Adjustable LED Flangeless Square Trim',
        status: '',
        quantity: 11
      },
      {
        sheetRow: 11,
        type: 'C1-A-TRIM',
        partNumber: 'E2SLB-OW',
        description: '2 inch Adjustable LED Flangeless Square Trim',
        status: '',
        quantity: 4
      }
    ]
  };

  var match = findBestTrackerMatch_(
    trackerInfo,
    {
      itemType: 'C1-2A-TRIM',
      itemName: 'E2SLB-OW',
      description: '2 inch Adjustable LED Flangeless Square Trim',
      qty: 11
    },
    {}
  );

  assertEqual(
    'Exact QBO Type disambiguates repeated part number',
    match.trackerRow.sheetRow,
    10
  );

  var ambiguous = findBestTrackerMatch_(
    {
      rows: [
        {
          sheetRow: 20,
          type: 'R1',
          partNumber: 'ABC-123',
          description: 'Same',
          status: '',
          quantity: 2
        },
        {
          sheetRow: 21,
          type: 'R1',
          partNumber: 'ABC-123',
          description: 'Same',
          status: '',
          quantity: 2
        }
      ]
    },
    {
      itemType: 'R1',
      itemName: 'ABC-123',
      description: 'Same',
      qty: 2
    },
    {}
  );

  assertTrue(
    'Tied best matches are flagged ambiguous',
    ambiguous && ambiguous.ambiguous === true
  );

  Logger.log('PUMA unit tests passed: ' + JSON.stringify(results));
  return results;
}
