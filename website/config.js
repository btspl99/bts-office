/* BTS Office website settings — edit these, nothing else needs to change between releases.
 * The download button reads updates/latest.yml (written by each release), so the version,
 * file name and size update themselves. The values below are only used if that file
 * can't be read (for example when you open index.html straight from your computer). */
window.BTS_SITE = {
  // Change this once the GitHub repository exists.
  github: 'https://github.com/btspl99/bts-office',
  supportEmail: 'btspl99@gmail.com',
  organisation: 'Bhoomiputra Tech Solutions Private Limited',
  organisationUrl: 'https://btspl.tech',
  fallback: { version: '1.2.0', file: 'BTS-Office-Setup-1.2.0.exe', size: 96284548, date: '2026-10-15' },
};
