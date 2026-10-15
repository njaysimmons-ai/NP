/**
 * Assigns clinicians to AM and PM groups for each day of the week on the
 * active weekly schedule tab.
 *
 * Nothing here depends on fixed column letters: the script finds every column
 * by its header text, so it keeps working when columns are added, removed or
 * moved on either tab.
 *
 * Clinician Data tab (header row 1):
 *   "Mon AM" ... "Fri PM"  ten availability checkboxes (TRUE = available).
 *                          The column just left of "Mon AM" holds the names.
 *   "Unallowed Groups"     e.g. "Group 1, Group 4" (matched against the
 *                          group labels shown on the weekly tab)
 *   "Permanent Groups"     e.g. "AM Group 2" or "PM Group 1"
 *   Group Used/Unused flags live in the name column further down, e.g.
 *   "AM Group 1" with "Used" or "Unused" in the cell to its right. A group
 *   marked "Unused" gets no clinician.
 *
 * Weekly schedule tab (header row 1):
 *   "AM"        first AM group column; AM groups run up to the "PM" column
 *   "PM"        first PM group column; PM groups run up to the "Tasks" column
 *   "Tasks"     names already on task that day (not placed in a group)
 *   "Out of Office"  (optional) names out that day (not placed or unplaced)
 *   "Unplaced"  where clinicians left without a group are written
 *   Day blocks start on rows 2 (Mon), 9 (Tue), 16 (Wed), 23 (Thu), 30 (Fri).
 *   Group labels such as "Group 2 (PHP)" sit 3 rows below each day's first
 *   row; groups whose label contains "IOP" are filled last.
 *
 * "AM Group N" / "PM Group N" in the flags and in Permanent Groups means the
 * Nth AM or PM column on the weekly tab, counting from the left.
 *
 * Everything lives inside assignClinicianGroups() so this file adds only that
 * one name to the Apps Script project and can't clash with other files.
 */
function assignClinicianGroups() {
  var CLINICIAN_SHEET_NAME = "Clinician Data";
  var EXCLUDED_SHEET_NAMES = ["Intern Task Schedule", CLINICIAN_SHEET_NAME];

  var DAY_CONFIGS = [
    { name: "Monday",    leaderRow: 2,  amKey: "monAM", pmKey: "monPM" },
    { name: "Tuesday",   leaderRow: 9,  amKey: "tueAM", pmKey: "tuePM" },
    { name: "Wednesday", leaderRow: 16, amKey: "wedAM", pmKey: "wedPM" },
    { name: "Thursday",  leaderRow: 23, amKey: "thuAM", pmKey: "thuPM" },
    { name: "Friday",    leaderRow: 30, amKey: "friAM", pmKey: "friPM" }
  ];
  var AVAILABILITY_KEYS = ["monAM", "monPM", "tueAM", "tuePM", "wedAM", "wedPM", "thuAM", "thuPM", "friAM", "friPM"];

  // Tasks and Unplaced cells used per day: the day's first row plus 5 more.
  // The 7th row of each block holds checkboxes and is left alone.
  var DAY_BLOCK_ROWS = 6;
  // Group labels ("Group 2 (PHP)") sit this many rows below a day's first row.
  var GROUP_LABEL_OFFSET = 3;

  var GROUP_FLAG_PATTERN = /^(am|pm)\s+group\s+(\d+)$/;

  function run() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ui = SpreadsheetApp.getUi();
    var currentSheet = ss.getActiveSheet();

    // 0. Do not run on the Intern Task Schedule or Clinician Data tabs
    var sheetName = currentSheet.getName();
    if (EXCLUDED_SHEET_NAMES.indexOf(sheetName) !== -1) {
      ui.alert("This script cannot be run on the '" + sheetName + "' tab. Please switch to a weekly schedule tab first.");
      return;
    }

    var layout = findScheduleLayout(currentSheet);
    if (layout.error) {
      ui.alert("Couldn't read the layout of '" + sheetName + "': " + layout.error +
        "\n\nRow 1 needs headers named AM, PM, Tasks and Unplaced (in that order, left to right).");
      return;
    }

    // 1. Previous tab (one to the left), used to avoid repeating last week's leader
    var sheets = ss.getSheets();
    var currentPos = currentSheet.getIndex() - 1; // 0-based
    var previousSheet = currentPos > 0 ? sheets[currentPos - 1] : null;
    var previousLayout = previousSheet ? findScheduleLayout(previousSheet) : null;
    var previousValues = (previousLayout && !previousLayout.error)
      ? previousSheet.getDataRange().getValues()
      : null;

    // 2. Clinician Data
    var staffSheet = ss.getSheetByName(CLINICIAN_SHEET_NAME);
    if (!staffSheet) {
      ui.alert("Please ensure you have a tab titled '" + CLINICIAN_SHEET_NAME + "'.");
      return;
    }
    var staff = readClinicianData(staffSheet);
    if (staff.error) {
      ui.alert(staff.error);
      return;
    }
    var clinicians = staff.clinicians;
    if (clinicians.length === 0) {
      ui.alert("No clinicians found on the '" + CLINICIAN_SHEET_NAME + "' tab.");
      return;
    }
    var groupActiveMap = staff.groupActiveMap;

    // Defaults to active when a group has no Used/Unused row.
    function isGroupActive(side, groupIndex) {
      var label = side.toLowerCase() + " group " + (groupIndex + 1);
      return groupActiveMap.hasOwnProperty(label) ? groupActiveMap[label] : true;
    }

    var amGroupCols = layout.amCols;
    var pmGroupCols = layout.pmCols;

    // "AM Group 2" -> { side: "AM", groupIndex: 1 }; null when blank or invalid
    function parsePermanentGroup(pgStr) {
      var match = String(pgStr || "").trim().toLowerCase().match(GROUP_FLAG_PATTERN);
      if (!match) return null;
      var side = match[1].toUpperCase();
      var idx = parseInt(match[2], 10) - 1;
      var count = side === "AM" ? amGroupCols.length : pmGroupCols.length;
      return (idx >= 0 && idx < count) ? { side: side, groupIndex: idx } : null;
    }
    clinicians.forEach(function(c) { c.permanent = parsePermanentGroup(c.permanentGroup); });

    var sheetValues = currentSheet.getDataRange().getValues();
    function cellAt(values, row, col) {
      var line = values[row - 1];
      return (line && col - 1 < line.length) ? line[col - 1] : "";
    }

    // 3. Weekly tracking
    var weeklyAMGroup = {}; // name key -> [AM group indexes used this week]
    var weeklyPMGroup = {};
    var weeklyAMCount = {}; // name key -> number of AM assignments this week
    var previousDayUnplaced = [];

    // 4. Process each day
    DAY_CONFIGS.forEach(function(dayConfig) {
      var leaderRow = dayConfig.leaderRow;
      var labelRow = leaderRow + GROUP_LABEL_OFFSET;

      function groupInfo(cols, side) {
        return cols.map(function(col, idx) {
          var label = String(cellAt(sheetValues, labelRow, col) || "").trim();
          var shortName = label.replace(/\(.*?\)/g, "").trim() || ("Group " + (idx + 1));
          var prevCol = (previousLayout && !previousLayout.error)
            ? (side === "AM" ? previousLayout.amCols : previousLayout.pmCols)[idx]
            : null;
          return {
            col: col,
            name: shortName,
            isIOP: /iop/i.test(label),
            active: isGroupActive(side, idx),
            prevKey: prevCol ? nameKey(cellAt(previousValues, leaderRow, prevCol)) : ""
          };
        });
      }
      var amGroups = groupInfo(amGroupCols, "AM");
      var pmGroups = groupInfo(pmGroupCols, "PM");

      // Names already listed under Tasks for this day are not placed in a group
      var onTaskKeys = [];
      for (var r = leaderRow; r < leaderRow + DAY_BLOCK_ROWS; r++) {
        var taskKey = nameKey(cellAt(sheetValues, r, layout.tasksCol));
        if (taskKey) onTaskKeys.push(taskKey);
      }

      // Names listed under Out of Office for this day are skipped entirely
      var outKeys = [];
      if (layout.outOfOfficeCol) {
        for (var r = leaderRow; r < leaderRow + DAY_BLOCK_ROWS; r++) {
          var outKey = nameKey(cellAt(sheetValues, r, layout.outOfOfficeCol));
          if (outKey) outKeys.push(outKey);
        }
      }
      function isAvailable(c, availKey) {
        return c[availKey] && outKeys.indexOf(c.key) === -1;
      }

      var usedToday = onTaskKeys.slice();
      var amAssignments = blankArray(amGroups.length);
      var pmAssignments = blankArray(pmGroups.length);

      // STEP 0: permanent-group clinicians first
      clinicians.forEach(function(c) {
        var pg = c.permanent;
        if (!pg) return;
        var canAM = isAvailable(c, dayConfig.amKey);
        var canPM = isAvailable(c, dayConfig.pmKey);
        if (!isGroupActive(pg.side, pg.groupIndex)) return;

        if (pg.side === "AM") {
          if (!canAM || amAssignments[pg.groupIndex] !== "") return;
          amAssignments[pg.groupIndex] = c.name;
          recordGroup(weeklyAMGroup, c.key, pg.groupIndex);
        } else {
          if (!canPM || pmAssignments[pg.groupIndex] !== "") return;
          pmAssignments[pg.groupIndex] = c.name;
          recordGroup(weeklyPMGroup, c.key, pg.groupIndex);
        }
        usedToday.push(c.key);
      });

      // Pools for everyone else, skipping anyone on task today
      var amOnlyPool = [];
      var pmOnlyPool = [];
      var bothPool = [];
      clinicians.forEach(function(c) {
        if (c.permanent) return;
        if (onTaskKeys.indexOf(c.key) !== -1) return;
        var canAM = isAvailable(c, dayConfig.amKey);
        var canPM = isAvailable(c, dayConfig.pmKey);
        if (canAM && canPM) bothPool.push(c);
        else if (canAM) amOnlyPool.push(c);
        else if (canPM) pmOnlyPool.push(c);
        // neither -> out of office
      });

      // Clinicians left unplaced yesterday go to the front
      function prioritizeUnplaced(pool) {
        return pool.sort(function(a, b) {
          var aPri = previousDayUnplaced.indexOf(a.key) !== -1 ? 0 : 1;
          var bPri = previousDayUnplaced.indexOf(b.key) !== -1 ? 0 : 1;
          return aPri - bPri;
        });
      }
      amOnlyPool = prioritizeUnplaced(shuffleArray(amOnlyPool));
      pmOnlyPool = prioritizeUnplaced(shuffleArray(pmOnlyPool));
      bothPool = prioritizeUnplaced(shuffleArray(bothPool));

      function fillGroups(groups, assignments, candidates, weeklyGroup, availKey) {
        fillOrder(groups).forEach(function(g) {
          var group = groups[g];
          if (!group.active) return;              // marked Unused
          if (assignments[g] !== "") return;      // taken by a permanent clinician
          for (var j = 0; j < candidates.length; j++) {
            var c = candidates[j];
            if (!isAvailable(c, availKey)) continue;
            if (usedToday.indexOf(c.key) !== -1) continue;
            if (c.key === group.prevKey) continue; // led this group last week
            if (isUnallowedGroup(c.unallowedGroups, group.name)) continue;
            if (weeklyGroup[c.key] && weeklyGroup[c.key].indexOf(g) !== -1) continue;
            assignments[g] = c.name;
            usedToday.push(c.key);
            candidates.splice(j, 1);
            return;
          }
        });
      }

      // STEP 1: AM groups (PHP first, IOP last)
      fillGroups(amGroups, amAssignments, amOnlyPool.concat(bothPool), weeklyAMGroup, dayConfig.amKey);

      // STEP 2: PM groups (PHP first, IOP last). PM-only clinicians first, then
      // clinicians available all day, most AM assignments this week first.
      var unusedBoth = bothPool.filter(function(c) { return usedToday.indexOf(c.key) === -1; });
      unusedBoth.sort(function(a, b) {
        return (weeklyAMCount[b.key] || 0) - (weeklyAMCount[a.key] || 0);
      });
      fillGroups(pmGroups, pmAssignments, pmOnlyPool.concat(unusedBoth), weeklyPMGroup, dayConfig.pmKey);

      // STEP 3: write the leader row in one call
      var allCols = amGroupCols.concat(pmGroupCols);
      var allAssignments = amAssignments.concat(pmAssignments);
      var firstCol = Math.min.apply(null, allCols);
      var lastCol = Math.max.apply(null, allCols);
      var rowRange = currentSheet.getRange(leaderRow, firstCol, 1, lastCol - firstCol + 1);
      var rowValues = rowRange.getValues();
      allCols.forEach(function(col, i) { rowValues[0][col - firstCol] = allAssignments[i]; });
      writeIgnoringValidation(rowRange, rowValues);

      // STEP 4: weekly tracking for non-permanent placements
      amAssignments.forEach(function(name, g) {
        if (name === "") return;
        var key = nameKey(name);
        if (isPermanent(clinicians, key)) return; // tracked in STEP 0
        recordGroup(weeklyAMGroup, key, g);
        weeklyAMCount[key] = (weeklyAMCount[key] || 0) + 1;
      });
      pmAssignments.forEach(function(name, g) {
        if (name === "") return;
        var key = nameKey(name);
        if (isPermanent(clinicians, key)) return;
        recordGroup(weeklyPMGroup, key, g);
      });

      // STEP 5: Unplaced column (available, not on task, not in a group)
      var unplaced = clinicians.filter(function(c) {
        if (!isAvailable(c, dayConfig.amKey) && !isAvailable(c, dayConfig.pmKey)) return false;
        if (onTaskKeys.indexOf(c.key) !== -1) return false;
        return usedToday.indexOf(c.key) === -1;
      });
      var unplacedValues = [];
      for (var u = 0; u < DAY_BLOCK_ROWS; u++) {
        unplacedValues.push([u < unplaced.length ? unplaced[u].name : ""]);
      }
      writeIgnoringValidation(
        currentSheet.getRange(leaderRow, layout.unplacedCol, DAY_BLOCK_ROWS, 1), unplacedValues);

      previousDayUnplaced = unplaced.map(function(c) { return c.key; });
    });

    ss.toast("Clinician groups assigned successfully!");
  }

  /**
   * Finds the group, Tasks and Unplaced columns from the headers in row 1.
   * Returns { amCols, pmCols, tasksCol, unplacedCol } or { error }.
   */
  function findScheduleLayout(sheet) {
    var lastCol = sheet.getLastColumn();
    if (lastCol < 1) return { error: "the tab is empty." };
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(normalizeHeader);

    function find(test) {
      for (var i = 0; i < headers.length; i++) if (test(headers[i])) return i + 1;
      return 0;
    }
    var amCol = find(function(h) { return h === "am"; });
    var pmCol = find(function(h) { return h === "pm"; });
    var tasksCol = find(function(h) { return h.indexOf("tasks") === 0; });
    var unplacedCol = find(function(h) { return h === "unplaced"; });
    var outOfOfficeCol = find(function(h) { return h === "out of office"; });

    var missing = [];
    if (!amCol) missing.push("AM");
    if (!pmCol) missing.push("PM");
    if (!tasksCol) missing.push("Tasks");
    if (!unplacedCol) missing.push("Unplaced");
    if (missing.length) return { error: "missing header(s) " + missing.join(", ") + "." };
    if (!(amCol < pmCol && pmCol < tasksCol)) {
      return { error: "the AM, PM and Tasks headers are out of order." };
    }
    return {
      amCols: range(amCol, pmCol - 1),
      pmCols: range(pmCol, tasksCol - 1),
      tasksCol: tasksCol,
      unplacedCol: unplacedCol,
      outOfOfficeCol: outOfOfficeCol
    };
  }

  /**
   * Reads clinicians and group Used/Unused flags from the Clinician Data tab.
   * Returns { clinicians, groupActiveMap } or { error }.
   */
  function readClinicianData(staffSheet) {
    var values = staffSheet.getDataRange().getValues();
    var headers = (values[0] || []).map(normalizeHeader);

    var monAMIdx = headers.indexOf("mon am");
    if (monAMIdx < 1) {
      return { error: "Couldn't find the 'Mon AM' header in row 1 of '" + CLINICIAN_SHEET_NAME +
        "' (clinician names must be in the column just left of it)." };
    }
    var nameIdx = monAMIdx - 1;
    var unallowedIdx = indexWhere(headers, function(h) { return h.indexOf("unallowed group") === 0; });
    var permanentIdx = indexWhere(headers, function(h) { return h.indexOf("permanent group") === 0; });

    var clinicians = [];
    var groupActiveMap = {};
    var seen = {};
    for (var r = 1; r < values.length; r++) {
      var row = values[r];
      var rawName = row[nameIdx];
      if (typeof rawName !== "string") continue; // blanks, stray checkboxes
      var name = rawName.trim();
      if (name === "") continue;

      // Group flag rows, e.g. "AM Group 1" | "Unused"
      var label = name.toLowerCase().replace(/\s+/g, " ");
      if (GROUP_FLAG_PATTERN.test(label)) {
        groupActiveMap[label] = String(row[nameIdx + 1] || "").trim().toLowerCase() !== "unused";
        continue;
      }

      var key = nameKey(name);
      if (seen[key]) continue;
      seen[key] = true;

      var c = {
        name: name,
        key: key,
        unallowedGroups: unallowedIdx >= 0 ? String(row[unallowedIdx] || "") : "",
        permanentGroup: permanentIdx >= 0 ? String(row[permanentIdx] || "").trim() : ""
      };
      AVAILABILITY_KEYS.forEach(function(k, i) {
        var v = row[monAMIdx + i];
        c[k] = v === true || String(v).trim().toUpperCase() === "TRUE";
      });
      clinicians.push(c);
    }
    return { clinicians: clinicians, groupActiveMap: groupActiveMap };
  }

  /**
   * Returns true if groupName appears in the clinician's unallowed groups string.
   * Handles "Group 1" matching "Group 1 (PHP)" in either direction, but never
   * "Group 1" matching "Group 10". Supports comma-separated values.
   */
  function isUnallowedGroup(unallowedStr, groupName) {
    if (!unallowedStr) return false;
    var search = stripGroupLabel(groupName);
    return String(unallowedStr).split(",").some(function(s) {
      var item = stripGroupLabel(s);
      return item !== "" && item === search;
    });
  }

  function stripGroupLabel(s) {
    return String(s || "").replace(/\(.*?\)/g, "").trim().toLowerCase().replace(/\s+/g, " ");
  }

  /** Fisher-Yates in-place shuffle. Returns the shuffled array. */
  function shuffleArray(array) {
    for (var i = array.length - 1; i > 0; i--) {
      var j = Math.floor(Math.random() * (i + 1));
      var tmp = array[i];
      array[i] = array[j];
      array[j] = tmp;
    }
    return array;
  }

  /**
   * Writes values to a range. Script writes aren't blocked by dropdown rules,
   * so the dropdowns are left untouched (rules such as "=Table2[Available]"
   * can't be re-applied from Apps Script once removed).
   */
  function writeIgnoringValidation(rng, values) {
    rng.setValues(values);
  }

  /** Group indexes in fill order: non-IOP groups first, IOP groups last. */
  function fillOrder(groups) {
    var order = [];
    groups.forEach(function(g, i) { if (!g.isIOP) order.push(i); });
    groups.forEach(function(g, i) { if (g.isIOP) order.push(i); });
    return order;
  }

  function recordGroup(tracker, key, groupIndex) {
    if (!tracker[key]) tracker[key] = [];
    tracker[key].push(groupIndex);
  }

  function isPermanent(clinicians, key) {
    for (var i = 0; i < clinicians.length; i++) {
      if (clinicians[i].key === key) return !!clinicians[i].permanent;
    }
    return false;
  }

  /** Case- and space-insensitive key for comparing names; "" for non-names. */
  function nameKey(v) {
    if (typeof v !== "string") return "";
    return v.trim().toLowerCase().replace(/\s+/g, " ");
  }

  function normalizeHeader(v) {
    return String(v === null || v === undefined ? "" : v).trim().toLowerCase().replace(/\s+/g, " ");
  }

  function indexWhere(arr, test) {
    for (var i = 0; i < arr.length; i++) if (test(arr[i])) return i;
    return -1;
  }

  function range(from, to) {
    var out = [];
    for (var i = from; i <= to; i++) out.push(i);
    return out;
  }

  function blankArray(n) {
    var out = [];
    for (var i = 0; i < n; i++) out.push("");
    return out;
  }

  run();
}
