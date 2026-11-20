// Site configuration. This is the only file the coordinator or web team needs to edit.
// Nothing here is secret: the client ID and tenant ID are public identifiers.
window.TRACKER_CONFIG = {
  // --- Microsoft 365 / Entra ID (see README "IT setup") ---
  // Leave clientId empty to run in demo mode with sample data.
  clientId: '',
  tenantId: '', // DeSales tenant ID or domain, e.g. 'desales.edu'
  // Must exactly match a "Single-page application" redirect URI on the app registration.
  // Defaults to this page's URL without query or hash.
  redirectUri: '',

  // SharePoint link to the Instruction folder (Library Staff – Instruction Services / Instruction).
  // Copy it with Share → "People in DeSales with the link" → Copy link, or the folder's address bar URL.
  instructionFolderUrl: '',

  // Phase 2 faculty request form (request-form.html) posts here.
  // A Power Automate "When an HTTP request is received" flow that adds a row to tbl_Requests.
  requestFlowUrl: '',

  // --- Lists (spec §3.3). Edit these to change dropdowns; no new release needed. ---
  lists: {
    librarians: [
      // Colors are fixed per person so charts and calendar match even when filters change.
      { name: 'Kathleen', color: 1 },
      { name: 'Lisa', color: 2 },
      { name: 'Michele', color: 3 },
    ],
    sessionTypes: ['One shot', 'Embedded', 'Workshop', 'Presentation', 'Orientation', 'Tour'],
    sessionFormats: ['Face to face', 'Synchronous', 'Asynchronous'],
    roomStatuses: ['Confirmed', 'Pending'],
    yesNo: ['Yes', 'No'],
    studentTypes: ['Undergraduate', 'Graduate', 'Adult Studies'],
    // Placeholder list: the coordinator confirms the official names (open question Q10).
    colleges: [
      'College of Arts & Sciences',
      'College of Business',
      'College of Health Professions',
      'School of Nursing',
    ],
    // Course prefix → College/School. Unmapped prefixes leave College/School blank and show as "Unassigned".
    prefixToCollege: {
      EN: 'College of Arts & Sciences', HI: 'College of Arts & Sciences', PS: 'College of Arts & Sciences',
      TH: 'College of Arts & Sciences', PL: 'College of Arts & Sciences', BI: 'College of Arts & Sciences',
      CJ: 'College of Arts & Sciences', CO: 'College of Arts & Sciences', ED: 'College of Arts & Sciences',
      BU: 'College of Business', AC: 'College of Business', MBA: 'College of Business', MK: 'College of Business',
      PA: 'College of Health Professions', DPT: 'College of Health Professions', CSD: 'College of Health Professions',
      NU: 'School of Nursing',
    },
    graduateThreshold: 500,
    locations: ['TTC', 'Gmb 153', 'Gmb 155', 'Gmb 210', 'Gmb 214', 'Gmb 232', 'Dlg 104', 'Dlg 105', 'Dlg 108', 'Dlg 109', 'Dlg 228',
      'Wills 106', 'Wills 223', 'PC Lab 2', 'Library', 'Air Products', 'Alcoves', 'TEAMS', 'Online', 'Video', 'TBA'],
    // Term date ranges, used to default the term from a date (open question Q4: confirm Winter).
    terms: [
      { name: 'FALL 26', from: '2026-08-24', to: '2026-12-18' },
      { name: 'WINTER 26', from: '2026-12-19', to: '2027-01-17' },
      { name: 'SPRING 27', from: '2027-01-18', to: '2027-05-14' },
      { name: 'SUMMER 27', from: '2027-05-15', to: '2027-08-23' },
    ],
  },

  // How often to re-read the workbook (spec §5): visible tab / background tab.
  syncSecondsVisible: 30,
  syncSecondsHidden: 120,
};
