/**
 * PUMA PROJECT RESOLVER — READ-ONLY FOUNDATION
 *
 * Purpose
 * -------
 * Resolve inconsistent human-entered project names against the live PUMA project
 * registry WITHOUT writing to project trackers.
 *
 * Design rules:
 *  1) Existing PUMA tracker/config data is authoritative.
 *  2) "Residence" and "Project" are optional for matching, not for display.
 *  3) Contractor/customer prefixes such as "Arcadia:Kinney Residence" are ignored
 *     for comparison, while the original text is preserved for audit.
 *  4) Known config mismatches + historically consistent Raw Tracker Overrides can
 *     act as aliases.
 *  5) Fuzzy/typo matching is NEVER an automatic confirmation by itself.
 *  6) Operational evidence (PO number, Type, Part Number) can confirm a project.
 *  7) Ambiguous/conflicting evidence returns REVIEW/CONFLICT instead of guessing.
 *  8) This file performs no sheet writes.
 */

const PUMA_PROJECT_RESOLVER = {
  TRACKER_SUFFIX: ' - Project Tracker',
  TRACKER_CONFIG_NAMES: ['Tracker config', 'Tracker Config'],
  OPEN_PROJECTS_NAMES: ['Open Projects'],
  RAW_PO_SHEET_NAME: 'RAW_PO_IMPORT',

  // Terms that commonly describe workflow/context rather than the project identity.
  CONTEXT_SUFFIX_PATTERNS: [
    /\s*\(\s*dock\s+pickup\s*\)\s*$/i,
    /\s*\(\s*housing\s+sample\s*\)\s*$/i,
    /\s*\(\s*sample\s*\)\s*$/i,
    /\s*-\s*owner\s*$/i,
    /\s*-\s*night\s+design\s*$/i,
    /\s*-\s*yoakum\s*$/i
  ],

  // Explicit business aliases that cannot be derived safely from punctuation/
  // Residence normalization alone. These are read-only resolver hints.
  EXPLICIT_ALIASES: {
    'maddux': 'Maddux Farmhouse',
    'maddux residence': 'Maddux Farmhouse',
    'gough': 'Gough Hastings Residence',
    'gough residence': 'Gough Hastings Residence',
    'stoney shore': 'Stony Shore',
    'stock': 'Misc Stock Purchases'
  },

  // Fuzzy results are suggestions only. They never auto-write or auto-confirm.
  FUZZY_SUGGESTION_MIN: 0.78,
  FUZZY_MARGIN_MIN: 0.12
};


/* ========================================================================== */
/* Public read-only entry points                                               */
/* ========================================================================== */

/**
 * Resolve one project string against the active PUMA workbook.
 * Returns an object; performs NO writes.
 */
function pumaResolveProjectReadOnly(inputName, evidence) {
  const ss = SpreadsheetApp.getActive();
  const registry = pumaBuildProjectRegistry_(ss);
  return pumaResolveProjectWithRegistry_(registry, inputName, evidence || {});
}


/**
 * Dry-run the currently unprocessed / unmatched RAW_PO_IMPORT project routing.
 *
 * Options:
 *   onlyUnmatched: true by default
 *   limit: optional max number of rows
 *
 * Returns an array and logs a compact summary. NO writes.
 */
function pumaAuditRawPoProjectResolution(options) {
  options = options || {};
  const onlyUnmatched = options.onlyUnmatched !== false;
  const limit = Number(options.limit || 0);

  const ss = SpreadsheetApp.getActive();
  const raw = ss.getSheetByName(PUMA_PROJECT_RESOLVER.RAW_PO_SHEET_NAME);
  if (!raw) throw new Error('RAW_PO_IMPORT sheet not found.');

  const values = raw.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0].map(v => String(v || '').trim());
  const col = pumaHeaderMap_(headers);

  const registry = pumaBuildProjectRegistry_(ss);
  const trackerEvidence = pumaBuildTrackerEvidenceIndex_(ss, registry);

  const out = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (pumaRowBlank_(row)) continue;

    const currentOutcome = pumaGetByHeader_(row, col, 'Outcome');
    if (
      onlyUnmatched &&
      currentOutcome &&
      !/Unmatched|Tracker sheet not found|Could not normalize/i.test(String(currentOutcome))
    ) {
      continue;
    }

    const rawRecord = {
      sheetRow: r + 1,
      projectRaw: pumaGetByHeader_(row, col, 'project_raw'),
      poNumber: pumaGetByHeader_(row, col, 'po_number'),
      itemName: pumaGetByHeader_(row, col, 'item_name'),
      itemType: pumaGetByHeader_(row, col, 'item_type'),
      description: pumaGetByHeader_(row, col, 'description'),
      qty: pumaGetByHeader_(row, col, 'qty'),
      unitCost: pumaGetByHeader_(row, col, 'unit_cost'),
      vendor: pumaGetByHeader_(row, col, 'vendor'),
      currentOutcome: currentOutcome,
      rawTrackerOverride: pumaGetByHeader_(row, col, 'Raw Tracker Override')
    };

    const evidence = pumaGatherRawPoEvidence_(rawRecord, registry, trackerEvidence);
    const resolution = pumaResolveProjectWithRegistry_(
      registry,
      rawRecord.projectRaw,
      evidence
    );

    out.push({
      rawRow: rawRecord.sheetRow,
      poNumber: String(rawRecord.poNumber || ''),
      projectRaw: String(rawRecord.projectRaw || ''),
      itemType: String(rawRecord.itemType || ''),
      itemName: String(rawRecord.itemName || ''),
      currentOutcome: String(rawRecord.currentOutcome || ''),
      currentOverride: String(rawRecord.rawTrackerOverride || ''),
      status: resolution.status,
      canonicalProject: resolution.canonicalProject || '',
      trackerSheetName: resolution.trackerSheetName || '',
      method: resolution.method || '',
      confidence: resolution.confidence || '',
      reasons: resolution.reasons || [],
      conflicts: resolution.conflicts || [],
      suggestions: resolution.suggestions || []
    });

    if (limit && out.length >= limit) break;
  }

  const summary = pumaSummarizeResolutionResults_(out);
  console.log('[PUMA PROJECT RESOLVER] RAW PO dry-run summary: ' + JSON.stringify(summary));
  return out;
}


/**
 * Audit the project registry itself for collisions / malformed tracker names.
 * NO writes.
 */
function pumaAuditProjectRegistryReadOnly() {
  const ss = SpreadsheetApp.getActive();
  const registry = pumaBuildProjectRegistry_(ss);

  const result = {
    totalProjects: registry.projects.length,
    malformedTrackers: registry.malformedTrackers.slice(),
    normalizedCollisions: [],
    configTrackerMismatches: registry.configTrackerMismatches.slice(),
    historicalOverrideConflicts: registry.historicalOverrideConflicts.slice()
  };

  Object.keys(registry.normalizedToProjectIds).forEach(key => {
    const ids = registry.normalizedToProjectIds[key];
    if (ids.length > 1) {
      result.normalizedCollisions.push({
        normalizedKey: key,
        projects: ids.map(id => registry.byId[id].canonicalName)
      });
    }
  });

  console.log('[PUMA PROJECT RESOLVER] Registry audit: ' + JSON.stringify(result));
  return result;
}


/* ========================================================================== */
/* Registry construction                                                       */
/* ========================================================================== */

function pumaBuildProjectRegistry_(ss) {
  const trackerSheets = ss.getSheets()
    .map(s => s.getName())
    .filter(n => / - Project Tracker$/i.test(n));

  const malformedTrackers = trackerSheets
    .filter(n => !String(n).replace(PUMA_PROJECT_RESOLVER.TRACKER_SUFFIX, '').trim());

  const config = pumaReadTrackerConfig_(ss);
  const openProjects = pumaReadOpenProjects_(ss);
  const openProjectsByKey = {};
  openProjects.forEach(rec => {
    const key = pumaNormalizeProjectKey_(rec.project);
    if (!key) return;
    if (!openProjectsByKey[key]) openProjectsByKey[key] = [];
    openProjectsByKey[key].push(rec);
  });

  const configByTracker = {};
  const configTrackerMismatches = [];

  config.forEach(rec => {
    if (!rec.trackerName) return;
    if (!configByTracker[rec.trackerName]) configByTracker[rec.trackerName] = [];
    configByTracker[rec.trackerName].push(rec);

    const trackerBase = rec.trackerName.replace(/ - Project Tracker$/i, '').trim();
    if (rec.project && trackerBase && rec.project !== trackerBase) {
      configTrackerMismatches.push({
        project: rec.project,
        trackerName: rec.trackerName,
        trackerBase: trackerBase
      });
    }
  });

  const projects = [];
  const byId = {};
  const trackerToProjectId = {};
  const normalizedToProjectIds = {};
  const aliasToProjectIds = {};

  trackerSheets.forEach((trackerName, i) => {
    const trackerBase = trackerName.replace(/ - Project Tracker$/i, '').trim();
    if (!trackerBase) return;

    const cfgRows = configByTracker[trackerName] || [];
    const configProject = cfgRows.length
      ? String(cfgRows[0].project || '').trim()
      : '';

    // Preserve the actual tracker destination separately from display/canonical name.
    // When config has a clearly maintained project label, prefer it; otherwise tracker base.
    const canonicalName = pumaChooseCanonicalName_(trackerBase, configProject);
    const id = 'TRACKER_' + String(i + 1).padStart(4, '0');

    const project = {
      id: id,
      canonicalName: canonicalName,
      trackerName: trackerName,
      trackerBase: trackerBase,
      aliases: [],
      configProjectNames: [...new Set(cfgRows.map(r => r.project).filter(Boolean))]
    };

    const aliases = new Set([
      canonicalName,
      trackerBase,
      configProject
    ].filter(Boolean));

    project.configProjectNames.forEach(v => aliases.add(v));

    // Automatically allow form without "Residence"/"Project".
    [...aliases].forEach(v => {
      const normalized = pumaNormalizeProjectKey_(v);
      if (normalized) aliases.add(normalized);
    });

    project.aliases = [...aliases];

    projects.push(project);
    byId[id] = project;
    trackerToProjectId[trackerName] = id;

    const keys = new Set();
    project.aliases.forEach(a => {
      const k = pumaNormalizeProjectKey_(a);
      if (k) keys.add(k);
    });

    keys.forEach(k => {
      if (!normalizedToProjectIds[k]) normalizedToProjectIds[k] = [];
      if (!normalizedToProjectIds[k].includes(id)) normalizedToProjectIds[k].push(id);

      if (!aliasToProjectIds[k]) aliasToProjectIds[k] = [];
      if (!aliasToProjectIds[k].includes(id)) aliasToProjectIds[k].push(id);
    });
  });

  // Add explicit aliases.
  Object.keys(PUMA_PROJECT_RESOLVER.EXPLICIT_ALIASES).forEach(alias => {
    const target = PUMA_PROJECT_RESOLVER.EXPLICIT_ALIASES[alias];
    const targetId = pumaFindProjectIdByName_(projects, target);
    if (!targetId) return;

    const key = pumaNormalizeProjectKey_(alias);
    if (!key) return;
    if (!aliasToProjectIds[key]) aliasToProjectIds[key] = [];
    if (!aliasToProjectIds[key].includes(targetId)) aliasToProjectIds[key].push(targetId);
  });

  // Add historically consistent Raw Tracker Overrides as aliases.
  const hist = pumaReadHistoricalOverrideAliases_(ss, trackerToProjectId);
  Object.keys(hist.consistentAliases).forEach(aliasKey => {
    const targetId = hist.consistentAliases[aliasKey];
    if (!aliasToProjectIds[aliasKey]) aliasToProjectIds[aliasKey] = [];
    if (!aliasToProjectIds[aliasKey].includes(targetId)) {
      aliasToProjectIds[aliasKey].push(targetId);
    }
  });

  return {
    projects: projects,
    byId: byId,
    trackerToProjectId: trackerToProjectId,
    normalizedToProjectIds: normalizedToProjectIds,
    aliasToProjectIds: aliasToProjectIds,
    openProjects: openProjects,
    openProjectsByKey: openProjectsByKey,
    malformedTrackers: malformedTrackers,
    configTrackerMismatches: pumaUniqueObjects_(configTrackerMismatches),
    historicalOverrideConflicts: hist.conflicts
  };
}


function pumaChooseCanonicalName_(trackerBase, configProject) {
  const t = String(trackerBase || '').trim();
  const c = String(configProject || '').trim();

  if (!c) return t;

  // Config names may contain workflow/client suffixes that are not the true project name.
  const cClean = pumaStripContextSuffixes_(c);
  const tKey = pumaNormalizeProjectKey_(t);
  const cKey = pumaNormalizeProjectKey_(cClean);

  // If they represent the same normalized job, prefer the more complete "Residence" form.
  if (tKey && tKey === cKey) {
    if (/\bresidence\b/i.test(cClean) && !/\bresidence\b/i.test(t)) return cClean;
    return t;
  }

  // Otherwise preserve tracker identity as canonical until an explicit alias/registry
  // decision says the business name differs (e.g. Maddux Farmhouse).
  return t;
}


function pumaReadOpenProjects_(ss) {
  const sh = pumaFindSheetByAnyName_(ss, PUMA_PROJECT_RESOLVER.OPEN_PROJECTS_NAMES);
  if (!sh || sh.getLastRow() < 2) return [];

  const values = sh.getRange(2, 1, sh.getLastRow() - 1, Math.min(2, sh.getLastColumn())).getValues();
  return values.map((row, i) => ({
    project: String(row[0] || '').trim(),
    folderId: String(row[1] || '').trim(),
    sourceRow: i + 2
  })).filter(r => r.project && r.folderId);
}


function pumaReadTrackerConfig_(ss) {
  const sh = pumaFindSheetByAnyName_(ss, PUMA_PROJECT_RESOLVER.TRACKER_CONFIG_NAMES);
  if (!sh) return [];

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0].map(v => String(v || '').trim());
  const map = pumaHeaderMap_(headers);

  return values.slice(1).map(row => ({
    enabled: pumaGetByHeader_(row, map, 'Enable?'),
    project: String(pumaGetByHeader_(row, map, 'Project') || '').trim(),
    trackerName: String(
      pumaGetByHeader_(row, map, 'Tracker Sheet Name') ||
      pumaGetByHeader_(row, map, 'Tracker Name') ||
      ''
    ).trim(),
    quoteName: String(pumaGetByHeader_(row, map, 'Quote Name') || '').trim(),
    quoteId: String(pumaGetByHeader_(row, map, 'Quote Sheet ID') || '').trim()
  })).filter(r => r.project || r.trackerName);
}


function pumaReadHistoricalOverrideAliases_(ss, trackerToProjectId) {
  const sh = ss.getSheetByName(PUMA_PROJECT_RESOLVER.RAW_PO_SHEET_NAME);
  if (!sh) return {consistentAliases: {}, conflicts: []};

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return {consistentAliases: {}, conflicts: []};

  const headers = values[0].map(v => String(v || '').trim());
  const map = pumaHeaderMap_(headers);

  const seen = {};
  for (let r = 1; r < values.length; r++) {
    const projectRaw = pumaGetByHeader_(values[r], map, 'project_raw');
    const override = String(
      pumaGetByHeader_(values[r], map, 'Raw Tracker Override') || ''
    ).trim();

    if (!projectRaw || !override || !trackerToProjectId[override]) continue;

    const key = pumaNormalizeProjectKey_(projectRaw);
    if (!key) continue;

    if (!seen[key]) seen[key] = {};
    const id = trackerToProjectId[override];
    seen[key][id] = (seen[key][id] || 0) + 1;
  }

  const consistentAliases = {};
  const conflicts = [];

  Object.keys(seen).forEach(key => {
    const ids = Object.keys(seen[key]);
    if (ids.length === 1) {
      consistentAliases[key] = ids[0];
    } else {
      conflicts.push({
        normalizedAlias: key,
        targets: ids.map(id => ({projectId: id, count: seen[key][id]}))
      });
    }
  });

  return {consistentAliases: consistentAliases, conflicts: conflicts};
}


/* ========================================================================== */
/* Project resolution                                                          */
/* ========================================================================== */

function pumaResolveProjectWithRegistry_(registry, inputName, evidence) {
  evidence = evidence || {};
  const original = String(inputName || '').trim();
  const key = pumaNormalizeProjectKey_(original);

  const result = {
    input: original,
    normalizedKey: key,
    status: 'UNKNOWN',
    confidence: 'NONE',
    method: '',
    canonicalProject: '',
    trackerSheetName: '',
    projectId: '',
    reasons: [],
    conflicts: [],
    suggestions: []
  };

  // Human-confirmed row override always wins IF it maps to an existing tracker.
  if (evidence.rawTrackerOverride) {
    const id = registry.trackerToProjectId[String(evidence.rawTrackerOverride).trim()];
    if (id) {
      return pumaResolvedResult_(
        result, registry.byId[id], 'CONFIRMED', 'HUMAN_OVERRIDE',
        ['Raw Tracker Override points to an existing PUMA tracker.']
      );
    }
  }

  // Exact/normalized aliases.
  if (key && registry.aliasToProjectIds[key]) {
    const ids = registry.aliasToProjectIds[key];
    if (ids.length === 1) {
      return pumaResolvedResult_(
        result, registry.byId[ids[0]], 'CONFIRMED', 'NORMALIZED_ALIAS',
        ['Incoming project name uniquely matches the PUMA project registry after conservative normalization.']
      );
    }

    if (ids.length > 1) {
      result.status = 'CONFLICT';
      result.confidence = 'NONE';
      result.method = 'AMBIGUOUS_ALIAS';
      result.conflicts.push('Normalized project name maps to more than one active tracker.');
      result.suggestions = ids.map(id => registry.byId[id].canonicalName);
      return result;
    }
  }

  // Explicit alias table.
  const explicitTarget = PUMA_PROJECT_RESOLVER.EXPLICIT_ALIASES[key];
  if (explicitTarget) {
    const id = pumaFindProjectIdByName_(registry.projects, explicitTarget);
    if (id) {
      return pumaResolvedResult_(
        result, registry.byId[id], 'CONFIRMED', 'EXPLICIT_ALIAS',
        ['Incoming project name matches an explicit business alias.']
      );
    }
  }

  // If this name belongs to the current Open Projects manifest but has no
  // existing tracker/alias, do not let generic PO/part evidence reroute it to
  // some other project. It needs tracker setup/reconciliation first.
  const openMatches = key ? (registry.openProjectsByKey[key] || []) : [];
  if (openMatches.length === 1) {
    result.status = 'REVIEW';
    result.confidence = 'NONE';
    result.method = 'OPEN_PROJECT_TRACKER_MISSING';
    result.canonicalProject = openMatches[0].project;
    result.reasons.push('Project exists in Open Projects but no existing tracker/alias resolves it.');
    result.suggestions = [openMatches[0].project];
    return result;
  }
  if (openMatches.length > 1) {
    result.status = 'CONFLICT';
    result.confidence = 'NONE';
    result.method = 'OPEN_PROJECT_COLLISION';
    result.conflicts.push('Multiple Open Projects rows normalize to the same project identity.');
    result.suggestions = openMatches.map(x => x.project);
    return result;
  }

  // Independent operational evidence.
  const evidenceIds = pumaEvidenceProjectIds_(evidence);
  if (evidenceIds.confirmed.length === 1 && evidenceIds.conflicts.length === 0) {
    const project = registry.byId[evidenceIds.confirmed[0]];
    if (project) {
      return pumaResolvedResult_(
        result, project, 'CONFIRMED', evidenceIds.method,
        evidenceIds.reasons
      );
    }
  }

  if (evidenceIds.conflicts.length) {
    result.status = 'CONFLICT';
    result.confidence = 'NONE';
    result.method = 'EVIDENCE_CONFLICT';
    result.conflicts = evidenceIds.conflicts;
    result.suggestions = evidenceIds.suggestions;
    return result;
  }

  // Fuzzy is suggestion only.
  const fuzzy = pumaFuzzySuggestions_(registry, key);
  result.suggestions = fuzzy.map(x => ({
    project: x.project.canonicalName,
    tracker: x.project.trackerName,
    similarity: x.score
  }));

  if (fuzzy.length) {
    result.status = 'REVIEW';
    result.confidence = 'SUGGESTION_ONLY';
    result.method = 'FUZZY_SUGGESTION';
    result.reasons.push('A similar project name exists, but fuzzy spelling alone is not allowed to confirm a write.');
  } else if (evidenceIds.suggestions && evidenceIds.suggestions.length) {
    result.status = 'REVIEW';
    result.confidence = 'SUPPORTING_EVIDENCE_ONLY';
    result.method = 'WEAK_OPERATIONAL_EVIDENCE';
    result.suggestions = evidenceIds.suggestions
      .map(id => registry.byId[id])
      .filter(Boolean)
      .map(p => ({project: p.canonicalName, tracker: p.trackerName}));
    result.reasons = result.reasons.concat(evidenceIds.reasons || []);
  }

  return result;
}


function pumaEvidenceProjectIds_(evidence) {
  const reasons = [];
  const confirmed = [];
  const conflicts = [];
  const suggestions = [];
  let method = '';

  // Only multi-field evidence can CONFIRM. PO-only and Part-only are supporting
  // signals because the same PO and same part can legitimately span projects.
  const strongSources = [
    ['PO_TYPE_PART_QTY', evidence.poTypePartQtyProjectIds],
    ['PO_PART_QTY', evidence.poPartQtyProjectIds],
    ['TYPE_PART_QTY', evidence.typePartQtyProjectIds]
  ];

  const weakSources = [
    ['PO_TYPE_PART', evidence.poTypePartProjectIds],
    ['PO_AND_ITEM', evidence.poItemProjectIds],
    ['TYPE_AND_PART', evidence.typePartProjectIds],
    ['PO_NUMBER', evidence.poProjectIds],
    ['PART_ONLY', evidence.partProjectIds]
  ];

  const strongSets = [];
  strongSources.forEach(([label, ids]) => {
    ids = [...new Set((ids || []).filter(Boolean))];
    if (!ids.length) return;
    if (ids.length === 1) {
      strongSets.push({label: label, id: ids[0]});
    } else {
      conflicts.push(label + ' evidence points to multiple PUMA projects.');
      ids.forEach(id => suggestions.push(id));
    }
  });

  const uniqueStrong = [...new Set(strongSets.map(x => x.id))];
  if (uniqueStrong.length === 1 && conflicts.length === 0) {
    confirmed.push(uniqueStrong[0]);
    const labels = strongSets.filter(x => x.id === uniqueStrong[0]).map(x => x.label);
    method = labels.includes('PO_TYPE_PART_QTY')
      ? 'PO_TYPE_PART_QTY_CROSSCHECK'
      : labels.includes('PO_PART_QTY')
        ? 'PO_PART_QTY_CROSSCHECK'
        : 'TYPE_PART_QTY_CROSSCHECK';
    reasons.push('Exact multi-field tracker evidence uniquely identifies one project: ' + labels.join(', ') + '.');
  } else if (uniqueStrong.length > 1) {
    conflicts.push('Strong evidence sources point to different PUMA projects.');
    uniqueStrong.forEach(id => suggestions.push(id));
  }

  // Weak evidence may only suggest/reinforce; it never confirms by itself.
  weakSources.forEach(([label, ids]) => {
    ids = [...new Set((ids || []).filter(Boolean))];
    if (ids.length === 1) {
      suggestions.push(ids[0]);
      reasons.push(label + ' is supporting evidence only.');
    } else if (ids.length > 1) {
      ids.forEach(id => suggestions.push(id));
    }
  });

  return {
    confirmed: confirmed,
    conflicts: conflicts,
    suggestions: [...new Set(suggestions)],
    method: method,
    reasons: reasons
  };
}

/* ========================================================================== */
/* RAW PO operational evidence                                                 */
/* ========================================================================== */

function pumaBuildTrackerEvidenceIndex_(ss, registry) {
  const index = {
    po: {},
    part: {},
    typePart: {},
    poItem: {},
    poTypePart: {},
    typePartQty: {},
    poPartQty: {},
    poTypePartQty: {}
  };

  registry.projects.forEach(project => {
    const sh = ss.getSheetByName(project.trackerName);
    if (!sh) return;

    const values = sh.getDataRange().getDisplayValues();
    if (!values.length) return;

    const headerRow = pumaFindTrackerHeaderRowIndexReadOnly_(values);
    if (headerRow === -1) return;

    const map = pumaHeaderMap_(values[headerRow]);

    for (let r = headerRow + 1; r < values.length; r++) {
      const row = values[r];
      if (pumaRowBlank_(row)) continue;

      const po = pumaNormalizePo_(pumaGetByHeader_(row, map, 'PO Number'));
      const part = pumaNormalizePart_(pumaGetByHeader_(row, map, 'Part Number'));
      const type = pumaNormalizeType_(pumaGetByHeader_(row, map, 'Type'));
      const qty = pumaNormalizeQty_(pumaGetByHeader_(row, map, 'Quantity'));

      if (po) pumaIndexPush_(index.po, po, project.id);
      if (part) pumaIndexPush_(index.part, part, project.id);
      if (type && part) pumaIndexPush_(index.typePart, type + '||' + part, project.id);
      if (po && part) pumaIndexPush_(index.poItem, po + '||' + part, project.id);
      if (po && type && part) pumaIndexPush_(index.poTypePart, po + '||' + type + '||' + part, project.id);
      if (type && part && qty) pumaIndexPush_(index.typePartQty, type + '||' + part + '||' + qty, project.id);
      if (po && part && qty) pumaIndexPush_(index.poPartQty, po + '||' + part + '||' + qty, project.id);
      if (po && type && part && qty) pumaIndexPush_(index.poTypePartQty, po + '||' + type + '||' + part + '||' + qty, project.id);
    }
  });

  // Deduplicate.
  Object.keys(index).forEach(group => {
    Object.keys(index[group]).forEach(k => {
      index[group][k] = [...new Set(index[group][k])];
    });
  });

  return index;
}


function pumaGatherRawPoEvidence_(rawRecord, registry, index) {
  const po = pumaNormalizePo_(rawRecord.poNumber);
  const part = pumaNormalizePart_(rawRecord.itemName);
  const type = pumaNormalizeType_(rawRecord.itemType);
  const qty = pumaNormalizeQty_(rawRecord.qty);

  return {
    rawTrackerOverride: rawRecord.rawTrackerOverride || '',
    poProjectIds: po ? (index.po[po] || []) : [],
    poItemProjectIds: (po && part) ? (index.poItem[po + '||' + part] || []) : [],
    poTypePartProjectIds: (po && type && part) ? (index.poTypePart[po + '||' + type + '||' + part] || []) : [],
    typePartProjectIds: (type && part) ? (index.typePart[type + '||' + part] || []) : [],
    partProjectIds: part ? (index.part[part] || []) : [],
    typePartQtyProjectIds: (type && part && qty) ? (index.typePartQty[type + '||' + part + '||' + qty] || []) : [],
    poPartQtyProjectIds: (po && part && qty) ? (index.poPartQty[po + '||' + part + '||' + qty] || []) : [],
    poTypePartQtyProjectIds: (po && type && part && qty) ? (index.poTypePartQty[po + '||' + type + '||' + part + '||' + qty] || []) : []
  };
}

/* ========================================================================== */
/* Normalization + fuzzy suggestion helpers                                    */
/* ========================================================================== */

function pumaNormalizeProjectKey_(value) {
  let s = String(value || '').trim();
  if (!s) return '';

  // QBO commonly supplies Parent/Customer:Project. Keep the leaf project.
  if (s.indexOf(':') !== -1) {
    const parts = s.split(':');
    s = parts[parts.length - 1].trim();
  }

  s = pumaStripContextSuffixes_(s);

  return s
    .toLowerCase()
    .replace(/\b(residence|project)\b/g, ' ')
    .replace(/[’']/g, '')
    .replace(/&/g, ' and ')
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}


function pumaStripContextSuffixes_(value) {
  let s = String(value || '').trim();
  PUMA_PROJECT_RESOLVER.CONTEXT_SUFFIX_PATTERNS.forEach(rx => {
    s = s.replace(rx, '').trim();
  });
  return s;
}


function pumaFuzzySuggestions_(registry, key) {
  if (!key) return [];

  const seen = {};
  const scored = [];

  registry.projects.forEach(project => {
    const candidates = [project.canonicalName, project.trackerBase]
      .concat(project.aliases || [])
      .map(pumaNormalizeProjectKey_)
      .filter(Boolean);

    let best = 0;
    candidates.forEach(c => {
      best = Math.max(best, pumaSimilarity_(key, c));
    });

    if (best >= PUMA_PROJECT_RESOLVER.FUZZY_SUGGESTION_MIN) {
      if (!seen[project.id] || best > seen[project.id]) {
        seen[project.id] = best;
      }
    }
  });

  Object.keys(seen).forEach(id => {
    scored.push({project: registry.byId[id], score: Number(seen[id].toFixed(3))});
  });

  scored.sort((a, b) => b.score - a.score);

  if (scored.length >= 2) {
    const margin = scored[0].score - scored[1].score;
    if (margin < PUMA_PROJECT_RESOLVER.FUZZY_MARGIN_MIN) {
      return scored.slice(0, 5);
    }
  }

  return scored.slice(0, 5);
}


function pumaSimilarity_(a, b) {
  a = String(a || '');
  b = String(b || '');
  if (!a && !b) return 1;
  if (!a || !b) return 0;
  const d = pumaLevenshtein_(a, b);
  return 1 - d / Math.max(a.length, b.length);
}


function pumaLevenshtein_(a, b) {
  const dp = [];
  for (let j = 0; j <= b.length; j++) dp[j] = j;

  for (let i = 1; i <= a.length; i++) {
    let prev = dp[0];
    dp[0] = i;

    for (let j = 1; j <= b.length; j++) {
      const old = dp[j];
      dp[j] = Math.min(
        dp[j] + 1,
        dp[j - 1] + 1,
        prev + (a.charAt(i - 1) === b.charAt(j - 1) ? 0 : 1)
      );
      prev = old;
    }
  }

  return dp[b.length];
}


/* ========================================================================== */
/* Generic helpers                                                             */
/* ========================================================================== */

function pumaResolvedResult_(base, project, status, method, reasons) {
  base.status = status;
  base.confidence = status === 'CONFIRMED' ? 'HIGH' : '';
  base.method = method;
  base.canonicalProject = project.canonicalName;
  base.trackerSheetName = project.trackerName;
  base.projectId = project.id;
  base.reasons = (base.reasons || []).concat(reasons || []);
  return base;
}


function pumaFindProjectIdByName_(projects, name) {
  const target = pumaNormalizeProjectKey_(name);
  const matches = projects.filter(p =>
    pumaNormalizeProjectKey_(p.canonicalName) === target ||
    pumaNormalizeProjectKey_(p.trackerBase) === target
  );
  return matches.length === 1 ? matches[0].id : '';
}


function pumaFindTrackerHeaderRowIndexReadOnly_(data) {
  const required = ['Source', 'Type', 'Part Number', 'Status', 'PO Number'];

  for (let r = 0; r < Math.min(data.length, 10); r++) {
    const map = pumaHeaderMap_(data[r]);
    if (required.every(h => map[pumaHeaderKey_(h)] != null)) return r;
  }

  return -1;
}


function pumaHeaderMap_(headers) {
  const map = {};
  (headers || []).forEach((h, i) => {
    const key = pumaHeaderKey_(h);
    if (key) map[key] = i;
  });
  return map;
}


function pumaHeaderKey_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '')
    .trim();
}


function pumaGetByHeader_(row, map, header) {
  const idx = map[pumaHeaderKey_(header)];
  return idx == null ? '' : row[idx];
}


function pumaFindSheetByAnyName_(ss, names) {
  for (let i = 0; i < names.length; i++) {
    const sh = ss.getSheetByName(names[i]);
    if (sh) return sh;
  }
  return null;
}


function pumaNormalizePo_(value) {
  return String(value || '').replace(/[^0-9]/g, '');
}


function pumaNormalizePart_(value) {
  return String(value || '')
    .toUpperCase()
    .replace(/\s+/g, '')
    .replace(/[^A-Z0-9\-]/g, '');
}


function pumaNormalizeType_(value) {
  return String(value || '')
    .toUpperCase()
    .replace(/\s+/g, '')
    .replace(/[^A-Z0-9.\-]/g, '');
}


function pumaNormalizeQty_(value) {
  if (value === '' || value == null) return '';
  const raw = String(value).replace(/[$,\s]/g, '').trim();
  if (!raw) return '';
  const n = Number(raw);
  if (!isNaN(n)) return String(Math.round(n * 10000) / 10000);
  return '';
}


function pumaIndexPush_(obj, key, value) {
  if (!key || !value) return;
  if (!obj[key]) obj[key] = [];
  obj[key].push(value);
}


function pumaRowBlank_(row) {
  return !(row || []).some(v => v !== '' && v != null);
}


function pumaUniqueObjects_(rows) {
  const seen = {};
  return rows.filter(r => {
    const k = JSON.stringify(r);
    if (seen[k]) return false;
    seen[k] = true;
    return true;
  });
}


function pumaSummarizeResolutionResults_(rows) {
  const out = {total: rows.length, confirmed: 0, review: 0, conflict: 0, unknown: 0};
  rows.forEach(r => {
    const k = String(r.status || '').toLowerCase();
    if (out[k] != null) out[k]++;
  });
  return out;
}
