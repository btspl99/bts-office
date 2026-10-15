/* BTS Office website settings — edit these, nothing else needs to change between releases.
 * The download button reads updates/latest.yml (written by each release), so the version,
 * file name and size update themselves. The values below are only used if that file
 * can't be read (for example when you open index.html straight from your computer). */
window.BTS_SITE = {
  // Change this once the GitHub repository exists.
  github: 'https://github.com/YOUR-GITHUB-ACCOUNT/bts-office',
  supportEmail: 'btspl99@gmail.com',
  organisation: 'Bhoomiputra Tech Solutions Pvt Ltd',
  fallback: { version: '1.1.0', file: 'BTS-Office-Setup-1.1.0.exe', size: 85005066, date: '2026-10-15' },
};
