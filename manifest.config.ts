import { defineManifest } from '@crxjs/vite-plugin'
import pkg from './package.json'

export default defineManifest({
  manifest_version: 3,
  name: 'Mail Organizer',
  version: pkg.version,
  description: pkg.description,

  permissions: ['storage', 'activeTab', 'scripting', 'downloads', 'alarms'],
  host_permissions: [
    'https://outlook.office.com/*',
    'https://outlook.office365.com/*',
    'https://outlook.cloud.microsoft/*',
    'https://api.anthropic.com/*',
  ],

  icons: {
    16: 'public/icons/icon-16.png',
    32: 'public/icons/icon-32.png',
    48: 'public/icons/icon-48.png',
    128: 'public/icons/icon-128.png',
  },

  action: {
    default_popup: 'src/popup/index.html',
    default_title: 'Mail Organizer',
    default_icon: {
      16: 'public/icons/icon-16.png',
      32: 'public/icons/icon-32.png',
      48: 'public/icons/icon-48.png',
    },
  },

  options_page: 'src/options/index.html',

  background: {
    service_worker: 'src/background/service-worker.ts',
    type: 'module',
  },

  content_scripts: [
    {
      matches: [
        'https://outlook.office.com/*',
        'https://outlook.office365.com/*',
        'https://outlook.cloud.microsoft/*',
      ],
      js: ['src/content/content.ts'],
      run_at: 'document_idle',
    },
    {
      // Injects the floating action button + iframe panel on OWA pages so
      // the lawyer can launch the classify flow without leaving the inbox
      // (parallel to clicking the Chrome toolbar icon, which still works).
      matches: [
        'https://outlook.office.com/*',
        'https://outlook.office365.com/*',
        'https://outlook.cloud.microsoft/*',
      ],
      js: ['src/content/owa-fab.ts'],
      run_at: 'document_idle',
    },
  ],

  // Allow OWA pages to load our popup UI inside an iframe via the FAB.
  // Without this, the iframe load would be blocked by Chrome's
  // chrome-extension://… isolation policy.
  web_accessible_resources: [
    {
      resources: ['src/popup/index.html', 'src/options/index.html', 'public/icons/*'],
      matches: [
        'https://outlook.office.com/*',
        'https://outlook.office365.com/*',
        'https://outlook.cloud.microsoft/*',
      ],
    },
  ],
})
