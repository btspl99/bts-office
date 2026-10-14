/* BTS Slides — starter templates shown in File ▸ New. Each builds a complete presentation. */
(function () {
  const SL = window.SL;
  const T = window.SLTemplates = {};
  const P = (t) => `<p>${t}</p>`;
  const UL = (items) => `<ul>${items.map(i => `<li>${i}</li>`).join('')}</ul>`;
  const OL = (items) => `<ol>${items.map(i => `<li>${i}</li>`).join('')}</ol>`;

  function deck(themeId, variant, slides) {
    const pres = { width: 1280, height: 720, theme: { id: themeId, variant: variant || 0 }, footer: { date: false, number: true, text: '', skipTitle: true }, slides: [] };
    for (const spec of slides) {
      const s = SL.newSlide(pres, spec.layout);
      const ph = s.elements.filter(e => e.ph);
      (spec.fill || []).forEach((html, i) => { if (ph[i] && html != null) { if (ph[i].type === 'image') return; ph[i].html = html; } });
      (spec.extra || []).forEach(e => s.elements.push(SL.newElement(e.type, e)));
      if (spec.transition) s.transition = spec.transition;
      if (spec.notes) s.notes = spec.notes;
      if (spec.animate) s.elements.filter(e => e.ph === 'body').forEach((e, i) => e.anim = { type: spec.animate, trigger: 'click', dur: 500, order: i });
      pres.slides.push(s);
    }
    return pres;
  }
  const fade = { type: 'fade', dur: 600 }, push = { type: 'push', dur: 700, dir: 'left' };
  const card = (x, y, w, h, title, body, acc) => [
    { type: 'shape', kind: 'roundRect', x, y, w, h, fill: 'bg2', stroke: 'none', shadow: true, adj: .08, html: '' },
    { type: 'shape', kind: 'ellipse', x: x + 28, y: y + 28, w: 56, h: 56, fill: acc, html: '' },
    { type: 'text', x: x + 24, y: y + 100, w: w - 48, h: 44, html: P('<b>' + title + '</b>'), size: 22, role: 'title' },
    { type: 'text', x: x + 24, y: y + 148, w: w - 48, h: h - 170, html: P(body), size: 15 },
  ];
  const stat = (x, y, num, label, acc) => [
    { type: 'text', x: x - 30, y, w: 320, h: 90, html: P(num), size: 50, bold: true, color: acc, role: 'title', align: 'center', wrap: false },
    { type: 'text', x: x - 30, y: y + 92, w: 320, h: 60, html: P(label), size: 16, align: 'center' },
  ];

  T.list = [
    {
      name: 'Class presentation', theme: 'sunrise', build: () => deck('sunrise', 0, [
        { layout: 'title', fill: [P('Photosynthesis'), P('How plants turn sunlight into food · Class 8 Science')], notes: 'Introduce yourself and the topic.' },
        { layout: 'titleContent', fill: [P('What is photosynthesis?'), UL(['The process plants use to make glucose', 'Happens in the chloroplasts of leaf cells', 'Needs sunlight, water and carbon dioxide', 'Releases oxygen as a by-product'])], animate: 'fade', transition: fade },
        { layout: 'twoContent', fill: [P('Inputs and outputs'), UL(['<b>Inputs</b>', 'Sunlight', 'Water (H₂O)', 'Carbon dioxide (CO₂)']), UL(['<b>Outputs</b>', 'Glucose (C₆H₁₂O₆)', 'Oxygen (O₂)'])], transition: fade },
        { layout: 'titleOnly', fill: [P('The equation')], transition: fade, extra: [{ type: 'shape', kind: 'roundRect', x: 140, y: 270, w: 1000, h: 170, fill: 'accent1', html: P('6CO₂ + 6H₂O + light → C₆H₁₂O₆ + 6O₂'), size: 36, bold: true }] },
        { layout: 'section', fill: [P('Thank you!'), P('Questions?')], transition: fade },
      ]),
    },
    {
      name: 'Project report', theme: 'slate', build: () => deck('slate', 0, [
        { layout: 'title', fill: [P('Smart Water Monitor'), P('Final-year project report · Department of Computer Engineering')] },
        { layout: 'titleContent', fill: [P('Problem statement'), UL(['Households waste water through unnoticed leaks', 'Manual meter reading is slow and error-prone', 'Goal: real-time usage tracking with alerts'])], transition: push },
        { layout: 'titleOnly', fill: [P('Timeline')], transition: push, extra: [
          ...['Research', 'Design', 'Build', 'Test'].flatMap((t, i) => [{ type: 'shape', kind: 'chevron', x: 90 + i * 280, y: 300, w: 270, h: 110, fill: 'accent' + (i + 1), html: P('<b>' + t + '</b>'), size: 22 }, { type: 'text', x: 110 + i * 280, y: 430, w: 230, h: 40, html: P(['Jan – Feb', 'Mar', 'Apr – May', 'June'][i]), align: 'center', size: 16 }])] },
        { layout: 'titleOnly', fill: [P('Results')], transition: push, extra: [{ type: 'table', x: 140, y: 190, w: 1000, h: 280, header: true, banded: true, size: 18, accent: 'accent1', colW: [.4, .3, .3], rows: [['Metric', 'Before', 'After'], ['Daily usage (L)', '540', '410'], ['Leaks detected', '0', '3'], ['Bill (₹ / month)', '820', '610']].map(r => r.map(c => ({ html: P(c) }))) }] },
        { layout: 'titleContent', fill: [P('Conclusion'), UL(['24% reduction in water use', 'Low-cost hardware (under ₹1,500)', 'Next: mobile app and solar power'])], transition: push },
      ]),
    },
    {
      name: 'Science fair', theme: 'geo', build: () => deck('geo', 1, [
        { layout: 'title', fill: [P('Does Music Help Plants Grow?'), P('An experiment by Team Green · District Science Fair')] },
        { layout: 'titleContent', fill: [P('Hypothesis'), P('If plants are exposed to soft music for two hours a day, then they will grow taller than plants kept in silence.')], transition: fade },
        { layout: 'titleOnly', fill: [P('Our method')], transition: fade, extra: [...card(80, 180, 350, 420, '1. Setup', 'Six identical pots of mung beans, same soil, water and sunlight.', 'accent1'), ...card(465, 180, 350, 420, '2. Test', 'Three pots heard music for 2 hours daily; three stayed in silence.', 'accent2'), ...card(850, 180, 350, 420, '3. Measure', 'We measured the height of each plant every day for 14 days.', 'accent3')] },
        { layout: 'titleOnly', fill: [P('Results')], transition: fade, extra: [{ type: 'chart', x: 140, y: 170, w: 1000, h: 480, chart: { type: 'bar', title: 'Average height (cm)', categories: ['Day 4', 'Day 8', 'Day 12', 'Day 14'], series: [{ name: 'With music', values: [3.1, 7.4, 11.8, 13.6] }, { name: 'Silence', values: [2.9, 6.6, 10.1, 11.5] }] } }] },
        { layout: 'section', fill: [P('Conclusion'), P('Plants with music grew about 18% taller — our hypothesis was supported.')], transition: fade },
      ]),
    },
    {
      name: 'Seminar', theme: 'paper', build: () => deck('paper', 0, [
        { layout: 'title', fill: [P('Renewable Energy in India'), P('Seminar · Department of Electrical Engineering')] },
        { layout: 'titleContent', fill: [P('Agenda'), OL(['Where we are today', 'Solar and wind growth', 'Challenges', 'The road to 2030'])], transition: fade },
        { layout: 'titleOnly', fill: [P('Key numbers')], transition: fade, extra: [...stat(110, 260, '190+ GW', 'Renewable capacity', 'accent1'), ...stat(510, 260, '4th', 'Largest installed capacity in the world', 'accent2'), ...stat(910, 260, '500 GW', 'Non-fossil target for 2030', 'accent3')] },
        { layout: 'comparison', fill: [P('Solar vs Wind'), P('Solar'), UL(['Falling panel costs', 'Rooftop schemes', 'Works across most states']), P('Wind'), UL(['Strong coastal sites', 'Good at night & monsoon', 'Offshore potential'])], transition: fade },
        { layout: 'section', fill: [P('Thank you'), P('Questions & discussion')], transition: fade },
      ]),
    },
    {
      name: 'Pitch deck', theme: 'midnight', build: () => deck('midnight', 0, [
        { layout: 'title', fill: [P('StudyBuddy'), P('Peer tutoring for every campus')] },
        { layout: 'titleContent', fill: [P('The problem'), UL(['Students struggle alone before exams', 'Private coaching is expensive', 'Seniors who could help have no easy way to'])], transition: { type: 'zoom', dur: 600 } },
        { layout: 'titleOnly', fill: [P('Our solution')], transition: { type: 'zoom', dur: 600 }, extra: [
          { type: 'icon', name: 'users', x: 160, y: 230, w: 120, h: 120, color: 'accent1' }, { type: 'text', x: 100, y: 370, w: 240, h: 90, html: P('<b>Match</b><br>with seniors who aced the subject'), align: 'center', size: 18 },
          { type: 'icon', name: 'calendar-check', x: 580, y: 230, w: 120, h: 120, color: 'accent2' }, { type: 'text', x: 520, y: 370, w: 240, h: 90, html: P('<b>Book</b><br>30-minute sessions in two taps'), align: 'center', size: 18 },
          { type: 'icon', name: 'trophy', x: 1000, y: 230, w: 120, h: 120, color: 'accent3' }, { type: 'text', x: 940, y: 370, w: 240, h: 90, html: P('<b>Earn</b><br>tutors get paid or earn credits'), align: 'center', size: 18 }] },
        { layout: 'titleOnly', fill: [P('Traction')], transition: { type: 'zoom', dur: 600 }, extra: [{ type: 'chart', x: 160, y: 170, w: 960, h: 470, chart: { type: 'line', title: 'Monthly sessions', categories: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'], series: [{ name: 'Sessions', values: [120, 260, 410, 780, 1150, 1720] }] } }] },
        { layout: 'section', fill: [P('Join us'), P('hello@studybuddy.example')], transition: { type: 'zoom', dur: 600 } },
      ]),
    },
  ];
})();
