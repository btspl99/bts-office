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
      (spec.fill || []).forEach((html, i) => { if (ph[i] && html) { if (ph[i].type === 'image') return; ph[i].html = html; } });
      if (spec.dropEmpty) s.elements = s.elements.filter(e => !(e.ph && e.type === 'text' && SL.isEmptyText(e)));
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
    /* ---------- 1.2 professional templates ---------- */
    {
      name: 'Business pitch', theme: 'corporate', build: () => deck('corporate', 0, [
        { layout: 'title', fill: [P('GreenGrid Energy'), P('Affordable rooftop solar for small businesses · Investor presentation 2026')], notes: 'Open with the one-line story: we make solar as easy as paying an electricity bill.' },
        { layout: 'titleContent', fill: [P('The problem'), UL(['Small shops pay ₹9–12 per unit for grid power', 'Rooftop solar needs ₹3–5 lakh upfront', 'Installers are fragmented and hard to trust'])], transition: fade, animate: 'fade', notes: 'Use one shopkeeper story to make it real.' },
        { layout: 'titleOnly', fill: [P('Our solution')], transition: fade, extra: [
          { type: 'icon', name: 'sun', x: 170, y: 220, w: 110, h: 110, color: 'accent1', alt: 'Sun icon' }, { type: 'text', x: 100, y: 345, w: 250, h: 110, html: P('<b>Zero upfront</b><br>Pay monthly from the savings'), align: 'center', size: 18 },
          { type: 'icon', name: 'smartphone', x: 585, y: 220, w: 110, h: 110, color: 'accent2', alt: 'Phone icon' }, { type: 'text', x: 515, y: 345, w: 250, h: 110, html: P('<b>App-managed</b><br>Live savings and service'), align: 'center', size: 18 },
          { type: 'icon', name: 'shield-check', x: 1000, y: 220, w: 110, h: 110, color: 'accent5', alt: 'Shield icon' }, { type: 'text', x: 930, y: 345, w: 250, h: 110, html: P('<b>10-year warranty</b><br>Certified installers only'), align: 'center', size: 18 }] },
        { layout: 'titleOnly', fill: [P('Market opportunity')], transition: fade, extra: [...stat(110, 250, '6.3 cr', 'Small businesses in India', 'accent1'), ...stat(510, 250, '₹1.2 L cr', 'Annual electricity spend', 'accent2'), ...stat(910, 250, '18%', 'Yearly growth of rooftop solar', 'accent5')] },
        { layout: 'titleOnly', fill: [P('Business model')], transition: fade, extra: [{ type: 'table', x: 140, y: 190, w: 1000, h: 300, header: true, banded: true, tstyle: 'light', size: 18, accent: 'accent1', colW: [.4, .3, .3], rows: [['Revenue stream', 'Price', 'Margin'], ['Solar subscription', '₹2,400 / month', '32%'], ['Maintenance plan', '₹299 / month', '55%'], ['Battery add-on', '₹900 / month', '28%']].map(r => r.map(c => ({ html: P(c) }))) }] },
        { layout: 'titleOnly', fill: [P('Traction')], transition: fade, extra: [{ type: 'chart', x: 160, y: 170, w: 960, h: 470, alt: 'Line chart: installed sites grew from 12 in Q1 to 310 in Q4', chart: { type: 'line', title: 'Installed sites', categories: ['Q1', 'Q2', 'Q3', 'Q4'], series: [{ name: 'Sites', values: [12, 48, 140, 310] }] } }] },
        { layout: 'section', fill: [P('The ask: ₹4 crore seed'), P('18 months of runway · 2,000 installations · 3 new cities')], transition: fade },
      ]),
    },
    {
      name: 'Thesis defense', theme: 'academic', build: () => deck('academic', 0, [
        { layout: 'title', fill: [P('Low-Cost Air Quality Sensing for Indian Cities'), P('M.Tech thesis defense · Priya Sharma · Guide: Dr. R. Iyer · Department of Environmental Engineering')], notes: 'Thank the committee. State the research question in one sentence.' },
        { layout: 'titleContent', fill: [P('Outline'), OL(['Motivation and research questions', 'Related work', 'Sensor design and calibration', 'Field deployment and results', 'Conclusions and future work'])], transition: fade },
        { layout: 'titleContent', fill: [P('Research questions'), OL(['Can ₹2,000 sensors match reference monitors within 15%?', 'How does humidity affect PM2.5 readings?', 'What sensor density captures street-level variation?'])], transition: fade, animate: 'fade' },
        { layout: 'titleOnly', fill: [P('Methodology')], transition: fade, extra: ['Design', 'Calibrate', 'Deploy', 'Analyse'].flatMap((t, i) => [{ type: 'shape', kind: 'chevron', x: 90 + i * 280, y: 290, w: 270, h: 110, fill: 'accent' + (i + 1), html: P('<b>' + t + '</b>'), size: 22 }, { type: 'text', x: 110 + i * 280, y: 420, w: 230, h: 70, html: P(['PMS5003 + ESP32 node', 'Co-location, 6 weeks', '24 sites, 4 months', 'Regression + kriging'][i]), align: 'center', size: 16 }]) },
        { layout: 'twoContent', fill: [P('Results'), UL(['Mean error 11.4% after calibration', 'Humidity above 75% inflates PM2.5 by 30%', '400 m spacing captures 90% of variance'])], dropEmpty: true, transition: fade, extra: [{ type: 'chart', x: 660, y: 170, w: 540, h: 440, alt: 'Column chart comparing sensor error before and after calibration across three seasons', chart: { type: 'bar', title: 'Error vs reference (%)', categories: ['Winter', 'Summer', 'Monsoon'], series: [{ name: 'Raw', values: [24, 19, 38] }, { name: 'Calibrated', values: [10, 9, 15] }] } }] },
        { layout: 'titleOnly', fill: [P('Comparison with reference monitors')], transition: fade, extra: [{ type: 'table', x: 140, y: 190, w: 1000, h: 280, header: true, banded: true, tstyle: 'grid', size: 18, accent: 'accent1', colW: [.34, .22, .22, .22], rows: [['Metric', 'Our sensor', 'Reference', 'Difference'], ['PM2.5 mean (µg/m³)', '61.2', '57.9', '+5.7%'], ['R²', '0.91', '1.00', '—'], ['Cost per node', '₹2,000', '₹8,00,000', '400× cheaper']].map(r => r.map(c => ({ html: P(c) }))) }] },
        { layout: 'titleContent', fill: [P('Conclusions and future work'), UL(['Low-cost nodes are accurate enough for neighbourhood maps', 'Humidity correction is essential in coastal cities', 'Next: solar-powered nodes and a public dashboard'])], transition: fade },
        { layout: 'section', fill: [P('Thank you'), P('Questions and discussion')], transition: fade },
      ]),
    },
    {
      name: 'Product launch', theme: 'darkpro', build: () => deck('darkpro', 0, [
        { layout: 'title', fill: [P('Introducing Nova Pen'), P('The smart pen that turns handwritten notes into searchable text')], notes: 'Pause after the product name. Let the visual land.' },
        { layout: 'section', fill: [P('Write naturally. Search instantly.'), P('Your notebook, now with Ctrl+F')], transition: { type: 'zoom', dur: 600 } },
        { layout: 'titleOnly', fill: [P('What makes it different')], transition: fade, extra: [
          { type: 'icon', name: 'scan-text', x: 170, y: 220, w: 110, h: 110, color: 'accent1', alt: 'Scan text icon' }, { type: 'text', x: 100, y: 345, w: 250, h: 110, html: P('<b>Live transcription</b><br>English, हिन्दी and 8 more'), align: 'center', size: 18 },
          { type: 'icon', name: 'battery-full', x: 585, y: 220, w: 110, h: 110, color: 'accent2', alt: 'Battery icon' }, { type: 'text', x: 515, y: 345, w: 250, h: 110, html: P('<b>14-day battery</b><br>USB-C fast charge'), align: 'center', size: 18 },
          { type: 'icon', name: 'cloud-off', x: 1000, y: 220, w: 110, h: 110, color: 'accent3', alt: 'Offline icon' }, { type: 'text', x: 930, y: 345, w: 250, h: 110, html: P('<b>Works offline</b><br>Private by design'), align: 'center', size: 18 }] },
        { layout: 'titleOnly', fill: [P('Specifications')], transition: fade, extra: [{ type: 'table', x: 140, y: 190, w: 1000, h: 300, header: true, banded: true, tstyle: 'dark', size: 18, accent: 'accent1', colW: [.4, .3, .3], rows: [['', 'Nova Pen', 'Nova Pen Pro'], ['Storage', '8 GB', '32 GB'], ['Battery', '14 days', '21 days'], ['Price', '₹4,999', '₹7,499']].map(r => r.map(c => ({ html: P(c) }))) }] },
        { layout: 'titleOnly', fill: [P('Launch timeline')], transition: fade, extra: ['Pre-orders', 'Reviews', 'Retail', 'Campus tour'].flatMap((t, i) => [{ type: 'shape', kind: 'homePlate', x: 90 + i * 280, y: 300, w: 270, h: 100, fill: { type: 'gradient', color: 'accent' + (i + 1), color2: 'accent' + (i + 1) + '/d35', angle: 90 }, html: P('<b>' + t + '</b>'), size: 22, shadow: true }, { type: 'text', x: 110 + i * 280, y: 420, w: 230, h: 40, html: P(['1 Nov', '15 Nov', '1 Dec', 'Jan 2027'][i]), align: 'center', size: 16 }]) },
        { layout: 'section', fill: [P('Available today'), P('novapen.example · ₹4,999')], transition: fade },
      ]),
    },
    {
      name: 'School project · हिन्दी', theme: 'bharat', build: () => deck('bharat', 0, [
        { layout: 'title', fill: [P('जल संरक्षण · Water Conservation'), P('कक्षा 7 विज्ञान परियोजना · Class 7 Science Project')], notes: 'नमस्ते! आज हम पानी बचाने के बारे में बात करेंगे।' },
        { layout: 'titleContent', fill: [P('पानी क्यों ज़रूरी है? · Why water matters'), UL(['पृथ्वी का केवल 3% पानी पीने योग्य है · Only 3% of Earth’s water is fresh', 'हर दिन एक व्यक्ति को 50 लीटर पानी चाहिए · A person needs 50 litres a day', 'कई गाँवों में गर्मियों में पानी की कमी · Many villages face summer shortages'])], transition: fade, animate: 'fade' },
        { layout: 'titleOnly', fill: [P('आँकड़े · Facts')], transition: fade, extra: [...stat(110, 250, '3%', 'मीठा पानी · Fresh water', 'accent1'), ...stat(510, 250, '18%', 'दुनिया की आबादी भारत में · World population in India', 'accent2'), ...stat(910, 250, '4%', 'दुनिया का पानी भारत में · World water in India', 'accent3')] },
        { layout: 'titleOnly', fill: [P('हम क्या कर सकते हैं? · What we can do')], transition: fade, extra: [
          { type: 'icon', name: 'droplets', x: 170, y: 220, w: 110, h: 110, color: 'accent3', alt: 'Water drops icon' }, { type: 'text', x: 90, y: 345, w: 270, h: 120, html: P('<b>नल बंद रखें</b><br>Turn off taps'), align: 'center', size: 20 },
          { type: 'icon', name: 'cloud-rain', x: 585, y: 220, w: 110, h: 110, color: 'accent2', alt: 'Rain icon' }, { type: 'text', x: 505, y: 345, w: 270, h: 120, html: P('<b>वर्षा जल संचयन</b><br>Harvest rainwater'), align: 'center', size: 20 },
          { type: 'icon', name: 'sprout', x: 1000, y: 220, w: 110, h: 110, color: 'accent1', alt: 'Plant icon' }, { type: 'text', x: 920, y: 345, w: 270, h: 120, html: P('<b>पेड़ लगाएँ</b><br>Plant trees'), align: 'center', size: 20 }] },
        { layout: 'twoContent', fill: [P('हमारा प्रयोग · Our experiment'), UL(['<b>हिन्दी</b>', 'छत से वर्षा जल इकट्ठा किया', 'एक हफ़्ते में 200 लीटर बचाया']), UL(['<b>English</b>', 'Collected rainwater from the roof', 'Saved 200 litres in one week'])], transition: fade },
        { layout: 'section', fill: [P('धन्यवाद · Thank you'), P('प्रश्न? · Questions?')], transition: fade },
      ]),
    },
  ];
})();
