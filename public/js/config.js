// Our own Firebase project (Google sign-in only). Firebase console → Project settings → Your apps → Web app.
export const firebaseConfig = {
  apiKey: 'AIzaSyCSpl4GTgkYvlSWFIqnkUABioPS-GZZOHE',
  authDomain: 'tusu-budget-pitching-feedback.firebaseapp.com',
  projectId: 'tusu-budget-pitching-feedback',
  appId: '1:905415164744:web:9b62ab5f93481fe6c8a97d',
};

// Apps Script → Deploy → Web app URL (ends in /exec). See README.md.
const LIVE_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbw5IN8b4kg-9k1XyDlC300QF20FXis1hBfoarULPnRoPFH7T5B7Bw33HevwXi86ETHS/exec';
// Local test runs only (localhost): a test deployment's /exec URL kept in this browser as localStorage 'scriptUrl'.
// The live site always uses the URL above.
const testUrl = () => { try { return ['localhost', '127.0.0.1'].includes(location.hostname) ? localStorage.getItem('scriptUrl') : null; } catch { return null; } };
export const SCRIPT_URL = testUrl() || LIVE_SCRIPT_URL;
