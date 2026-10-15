/* BTS Docs — starter templates for File ▸ New. */
(function () {
  const D = window.Docs;
  const today = () => new Date().toLocaleDateString('en-IN', { day: 'numeric', month: 'long', year: 'numeric' });
  const lines = (n, cls = '') => Array.from({ length: n }, (_, i) => `<div class="ln ${i % 3 === 2 ? 's' : i % 4 === 1 ? 'm' : ''} ${cls}"></div>`).join('');

  D.TEMPLATES = [
    {
      name: 'Assignment cover page',
      thumbHTML: `<div class="tthumb" style="text-align:center;padding-top:26px"><div style="height:16px;width:16px;border-radius:50%;background:#3C4FE0;margin:0 auto 10px"></div><div class="t1" style="font-size:10px">Assignment</div><div style="font-size:5px;color:#5a6275;margin-bottom:22px">Subject name</div>${lines(3)}<div style="margin-top:24px">${lines(4)}</div></div>`,
      apply() {
        D.page = D.defaultPage(); D.design = Object.assign(D.defaultDesign(), { border: { style: 'double', color: '#1f2a6b', width: 4 }, hdrFirst: false });
        D.setContent(`
<p style="text-align:center;margin-top:36pt"><span style="font-size:14pt;color:#5a6275;letter-spacing:2px">YOUR COLLEGE / SCHOOL NAME</span></p>
<p style="text-align:center"><span style="font-size:10pt;color:#7a8499">Affiliated to University of Mumbai · Department of Computer Engineering</span></p>
<p style="text-align:center;margin-top:60pt"><span style="font-size:12pt;color:#3C4FE0;letter-spacing:3px">ASSIGNMENT NO. 1</span></p>
<p class="title" style="text-align:center">Title of the Assignment</p>
<p class="subtitle" style="text-align:center">Subject: Data Structures and Algorithms (CSC303)</p>
<p style="text-align:center;margin-top:70pt"><b>Submitted by</b></p>
<table class="tbl t-minimal" style="width:60%;margin:0 auto 8pt"><colgroup><col style="width:45%"><col style="width:55%"></colgroup><tbody>
<tr><td><p>Name</p></td><td><p>Your full name</p></td></tr>
<tr><td><p>Roll number</p></td><td><p>00</p></td></tr>
<tr><td><p>Class / Division</p></td><td><p>SE Computer – A</p></td></tr>
<tr><td><p>Academic year</p></td><td><p>2026–27</p></td></tr>
</tbody></table>
<p style="text-align:center;margin-top:30pt"><b>Submitted to</b></p>
<p style="text-align:center">Prof. Faculty Name</p>
<p style="text-align:center;margin-top:40pt"><span style="color:#7a8499">Date of submission: ${today()}</span></p>
<hr class="page-break">
<h1>Aim</h1><p>Write the aim of the assignment here.</p>
<h1>Theory</h1><p>Explain the concepts in your own words. Use Heading 2 for sub-topics so they appear in the navigation pane.</p>
<h1>Conclusion</h1><p>Summarise what you learned.</p>`, { ftr: '<p style="text-align:center"><span style="font-size:9pt;color:#7a8499">Page <span class="fld" data-fld="page" contenteditable="false">1</span></span></p>' });
      },
    },
    {
      name: 'Resume',
      thumbHTML: `<div class="tthumb"><div class="t1">Your Name</div><div style="font-size:5px;color:#3C4FE0;margin-bottom:6px">email · phone · city</div><div class="t2">Education</div>${lines(3)}<div class="t2">Projects</div>${lines(4)}<div class="t2">Skills</div>${lines(2)}</div>`,
      apply() {
        D.page = Object.assign(D.defaultPage(), { m: { t: 0.7, b: 0.7, l: 0.8, r: 0.8 } }); D.design = Object.assign(D.defaultDesign(), { theme: 'bts' });
        D.setContent(`
<p class="title" style="margin-bottom:2pt">Aarav Deshmukh</p>
<p style="margin-bottom:10pt"><span style="color:#3C4FE0">aarav.deshmukh@email.com</span> · +91 98765 43210 · Pune, Maharashtra · <a href="https://linkedin.com/in/your-profile">linkedin.com/in/your-profile</a></p>
<p style="border-bottom:1px solid #3c4250;padding:2pt 4pt">Final-year computer engineering student who enjoys building practical software. Looking for an internship in web or backend development.</p>
<h2>Education</h2>
<p style="margin-bottom:0"><b>B.E. Computer Engineering</b> — Savitribai Phule Pune University<span class="tab">\t</span>2023 – 2027</p>
<p>CGPA: 8.6 / 10 (till Sem V)</p>
<p style="margin-bottom:0"><b>HSC (Science)</b> — Maharashtra State Board<span class="tab">\t</span>2023</p>
<p>Percentage: 88%</p>
<h2>Projects</h2>
<p style="margin-bottom:0"><b>Smart Attendance System</b> · Python, OpenCV, Flask</p>
<ul><li>Face-recognition attendance for classrooms; reduced roll-call time by 90%.</li><li>Deployed for 3 classes with 180 students.</li></ul>
<p style="margin-bottom:0"><b>College Events Portal</b> · React, Node.js, MongoDB</p>
<ul><li>Event registration and certificates for 2,000+ students.</li></ul>
<h2>Skills</h2>
<table class="tbl t-minimal t-noheader no-borders"><colgroup><col style="width:28%"><col style="width:72%"></colgroup><tbody>
<tr><td><p><b>Languages</b></p></td><td><p>C, C++, Java, Python, JavaScript, SQL</p></td></tr>
<tr><td><p><b>Tools</b></p></td><td><p>Git, Linux, VS Code, Docker, Figma</p></td></tr>
<tr><td><p><b>Languages spoken</b></p></td><td><p>English, Hindi, Marathi</p></td></tr>
</tbody></table>
<h2>Achievements</h2>
<ul><li>Winner, Smart India Hackathon (internal round), 2025.</li><li>Google Developer Student Club — core team member.</li></ul>`);
      },
    },
    {
      name: 'Formal letter',
      thumbHTML: `<div class="tthumb"><div style="text-align:right">${lines(3)}</div><div style="margin:10px 0">${lines(2)}</div><div class="t2" style="color:#333;font-size:6px">Subject: …</div>${lines(8)}<div style="margin-top:10px;width:40%">${lines(2)}</div></div>`,
      apply() {
        D.page = D.defaultPage(); D.design = D.defaultDesign();
        D.setContent(`
<p class="nospace">Your Name</p><p class="nospace">Flat 12, Shanti Niwas, Station Road</p><p class="nospace">Bhiwandi, Maharashtra – 421302</p><p class="nospace">your.email@example.com</p>
<p style="margin-top:12pt">${today()}</p>
<p class="nospace">To,</p><p class="nospace">The Principal</p><p class="nospace">ABC College of Engineering</p><p class="nospace">Thane, Maharashtra – 400601</p>
<p style="margin-top:12pt"><b>Subject: Request for a bonafide certificate</b></p>
<p>Respected Sir/Madam,</p>
<p style="text-align:justify">I am a student of SE Computer Engineering (Roll No. 24) in your college. I request you to kindly issue me a bonafide certificate, which is required for applying to a scholarship under the state government scheme.</p>
<p style="text-align:justify">I have attached a copy of my college ID card and the fee receipt for your reference. I would be grateful if the certificate could be issued at the earliest.</p>
<p>Thanking you.</p>
<p class="nospace" style="margin-top:12pt">Yours faithfully,</p><p class="nospace" style="margin-top:28pt">(Signature)</p><p class="nospace">Your Name</p>`);
      },
    },
    {
      name: 'Project report',
      thumbHTML: `<div class="tthumb"><div class="bar"></div><div class="t1">Project Report</div><div style="font-size:5px;color:#5a6275;margin-bottom:8px">Mini project · 2026</div><div class="t2">Contents</div>${lines(3)}<div class="t2">1. Introduction</div>${lines(5)}</div>`,
      apply() {
        D.page = D.defaultPage(); D.design = Object.assign(D.defaultDesign(), { theme: 'office' });
        D.setContent(`
<p class="title">Library Management System</p>
<p class="subtitle">Mini project report · Department of Computer Engineering · ${new Date().getFullYear()}</p>
<div class="toc" contenteditable="false"></div>
<hr class="page-break">
<h1>1. Introduction</h1>
<p style="text-align:justify">This report describes the design and implementation of a library management system for a college library. It covers the problem statement, objectives, system design, implementation and results.</p>
<h2>1.1 Problem statement</h2>
<p style="text-align:justify">Manual issue/return registers are slow, error-prone and make it difficult to track overdue books and fines.</p>
<h2>1.2 Objectives</h2>
<ol><li>Digitise the catalogue of books.</li><li>Automate issue, return and fine calculation.</li><li>Provide reports for librarians.</li></ol>
<h1>2. System design</h1>
<p>The system follows a three-tier architecture.</p>
<table class="tbl t-banded"><colgroup><col style="width:30%"><col style="width:70%"></colgroup><tbody>
<tr><td><p>Layer</p></td><td><p>Technology</p></td></tr>
<tr><td><p>Presentation</p></td><td><p>HTML, CSS, JavaScript</p></td></tr>
<tr><td><p>Application</p></td><td><p>Node.js with Express</p></td></tr>
<tr><td><p>Data</p></td><td><p>MySQL 8</p></td></tr>
</tbody></table>
<h1>3. Results</h1>
<p>Write your results here. Insert screenshots with Insert ▸ Picture and add captions from References ▸ Insert caption.</p>
<h1>4. Conclusion</h1>
<p>Summarise the outcome and possible future work.</p>
<h1>References</h1>
<ol><li>Sommerville, I. <i>Software Engineering</i>, 10th ed., Pearson, 2015.</li></ol>`, { hdr: '<p style="text-align:right"><span style="font-size:9pt;color:#7a8499">Library Management System</span></p>', ftr: '<p style="text-align:center"><span style="font-size:9pt">Page <span class="fld" data-fld="page" contenteditable="false">1</span> of <span class="fld" data-fld="pages" contenteditable="false">1</span></span></p>' });
        setTimeout(() => { D.updateTOC(true); D.clearHistory(); }, 150);
      },
    },
    {
      name: 'Class notes',
      thumbHTML: `<div class="tthumb"><div class="t1" style="font-size:9px">Unit 3 · Notes</div><div style="font-size:5px;color:#5a6275;margin-bottom:6px">Physics</div><div class="t2">Key ideas</div>${lines(2)}<div style="background:#eef0fb;padding:3px;margin:4px 0">${lines(2)}</div><div class="t2">Formulas</div>${lines(4)}</div>`,
      apply() {
        D.page = D.defaultPage(); D.design = Object.assign(D.defaultDesign(), { theme: 'ocean' });
        D.setContent(`
<p class="title" style="font-size:24pt">Unit 3 · Laws of Motion</p>
<p class="subtitle">Physics · Lecture notes · ${today()}</p>
<h2>Key ideas</h2>
<ul><li><b>Inertia</b> — a body continues in its state of rest or uniform motion unless acted upon by a net external force.</li><li><b>Force</b> causes a change in momentum: F = dp/dt.</li><li>Every action has an <b>equal and opposite</b> reaction.</li></ul>
<p style="background-color:#e8f3f4;border-left:3px solid #0E7C86;padding:2pt 4pt"><b>Remember:</b> Newton’s laws are valid only in inertial frames of reference.</p>
<h2>Formulas</h2>
<table class="tbl t-grid-accent"><colgroup><col style="width:40%"><col style="width:60%"></colgroup><tbody>
<tr><td><p>Quantity</p></td><td><p>Formula</p></td></tr>
<tr><td><p>Momentum</p></td><td><p>p = mv</p></td></tr>
<tr><td><p>Newton’s second law</p></td><td><p>F = ma</p></td></tr>
<tr><td><p>Impulse</p></td><td><p>J = FΔt = Δp</p></td></tr>
</tbody></table>
<h2>Questions to revise</h2>
<ol><li>Why do passengers lurch forward when a bus brakes suddenly?</li><li>A 2 kg ball moving at 5 m/s is stopped in 0.1 s. Find the force.</li></ol>
<h2>My doubts</h2><p><br></p>`);
      },
    },
    {
      name: 'Business proposal',
      thumbHTML: `<div class="tthumb" style="font-family:Inter,sans-serif"><div style="height:34px;background:#1F3A5F;margin:-12px -11px 8px;padding:9px 11px;color:#fff"><div style="font-size:9px;font-weight:700">Project Proposal</div><div style="font-size:4.5px;opacity:.8">Prepared for Sahyadri Agro Co-op</div></div><div class="t2" style="color:#1F3A5F;font-weight:700">Executive summary</div>${lines(4)}<div class="t2" style="color:#1F3A5F;font-weight:700">Pricing</div><div style="border-top:4px solid #1F3A5F;margin-top:3px">${lines(3)}</div></div>`,
      apply() {
        D.page = D.defaultPage(); D.design = Object.assign(D.defaultDesign(), { styleSet: 'corporate', accent: '#1F3A5F', dark: '#14213D' });
        BTS.setProps({ title: 'Project proposal — Smart irrigation pilot', subject: 'Business proposal', category: 'Proposal' });
        D.setContent(`
<p class="subtitle" style="margin-top:60pt">PROPOSAL · ${new Date().getFullYear()}</p>
<p class="title">Smart Irrigation Pilot for Sahyadri Agro Co-operative</p>
<p>Prepared by <b>Your Company Name</b> for the Board of Directors, Sahyadri Agro Co-operative Society, Nashik.</p>
<table class="tbl t-minimal" style="width:70%"><colgroup><col style="width:40%"><col style="width:60%"></colgroup><tbody>
<tr><td><p>Date</p></td><td><p>${today()}</p></td></tr><tr><td><p>Contact</p></td><td><p>Your Name · +91 98765 43210</p></td></tr><tr><td><p>Valid until</p></td><td><p>30 days from the date above</p></td></tr></tbody></table>
<hr class="page-break">
<h1>Executive summary</h1>
<p style="text-align:justify">We propose a six-month pilot of soil-moisture based drip irrigation on 40 acres of member farms. The pilot is expected to <span class="cs-strong">reduce water use by 30–35%</span> and raise grape yields by about 8%, paying for itself within two seasons.</p>
<h1>Objectives</h1>
<ul><li>Install 60 wireless soil-moisture sensors across 12 member farms.</li><li>Give every farmer a Marathi and English mobile dashboard.</li><li>Measure water savings and yield against control plots.</li></ul>
<h1>Scope and timeline</h1>
<table class="tbl t-banded"><colgroup><col style="width:22%"><col style="width:58%"><col style="width:20%"></colgroup><tbody>
<tr><td><p>Phase</p></td><td><p>Work</p></td><td><p>Duration</p></td></tr>
<tr><td><p>1. Survey</p></td><td><p>Soil testing, farm mapping, farmer meetings</p></td><td><p>3 weeks</p></td></tr>
<tr><td><p>2. Installation</p></td><td><p>Sensors, controllers, gateway and training</p></td><td><p>5 weeks</p></td></tr>
<tr><td><p>3. Monitoring</p></td><td><p>Weekly reports and support</p></td><td><p>4 months</p></td></tr></tbody></table>
<h1>Investment</h1>
<table class="tbl t-light"><colgroup><col style="width:70%"><col style="width:30%"></colgroup><tbody>
<tr><td><p>Item</p></td><td><p style="text-align:right">Amount (₹)</p></td></tr>
<tr><td><p>Sensors and controllers (60 units)</p></td><td><p style="text-align:right">4,20,000</p></td></tr>
<tr><td><p>Installation and training</p></td><td><p style="text-align:right">1,10,000</p></td></tr>
<tr><td><p>Monitoring and support (4 months)</p></td><td><p style="text-align:right">80,000</p></td></tr>
<tr><td><p><b>Total</b></p></td><td><p style="text-align:right"><b>6,10,000</b></p></td></tr></tbody></table>
<blockquote class="intense"><p>Next step: approve the pilot at the next board meeting so installation can start before the rabi season.</p></blockquote>
<h1>About us</h1><p>Write two or three sentences about your team, past projects and why you are the right partner.</p>`, { ftr: '<p style="text-align:right"><span style="font-size:9pt;color:#55627A">Smart Irrigation Pilot · Page <span class="fld" data-fld="page" contenteditable="false">1</span></span></p>' });
      },
    },
    {
      name: 'Research paper (APA)',
      thumbHTML: `<div class="tthumb" style="font-family:'EB Garamond',serif;text-align:center"><div style="font-size:8px;font-weight:700;margin:14px 0 3px">Climate Variability and Kharif Yields</div><div style="font-size:5px;margin-bottom:8px">Student Name · College</div><div style="font-size:6px;font-weight:700">Abstract</div>${lines(4)}<div style="font-size:6px;font-weight:700;margin-top:5px">References</div><div style="text-align:left">${lines(3)}</div></div>`,
      apply() {
        D.page = D.defaultPage(); D.design = Object.assign(D.defaultDesign(), { styleSet: 'academic', accent: '#222222', dark: '#111111', biblio: { style: 'apa', sources: [
          { id: 'sr1', type: 'journal', authors: 'Guntukula, Raju', title: 'Assessing the impact of climate change on Indian agriculture: Evidence from major crop yields', journal: 'Journal of Public Affairs', year: '2020', volume: '20', issue: '1', pages: 'e2040', doi: '10.1002/pa.2040' },
          { id: 'sr2', type: 'book', authors: 'Aggarwal, P. K.', title: 'Global climate change and Indian agriculture: Case studies from the ICAR network project', year: '2009', publisher: 'Indian Council of Agricultural Research', city: 'New Delhi' },
          { id: 'sr3', type: 'website', authors: 'India Meteorological Department', title: 'Statement on climate of India during 2023', journal: 'IMD', year: '2024', url: 'https://mausam.imd.gov.in/', accessed: '12 October 2026' },
        ] } });
        BTS.setProps({ title: 'Climate variability and kharif crop yields in Maharashtra', subject: 'Research paper', category: 'Research', keywords: 'climate, monsoon, agriculture, Maharashtra' });
        D.setContent(`
<p class="title">Climate Variability and Kharif Crop Yields in Maharashtra</p>
<p class="subtitle">Student Name<br>Department of Environmental Science, Your College<br>Course code: EVS-301 · Instructor: Prof. Name · ${today()}</p>
<h1>Abstract</h1>
<p style="text-align:justify">This paper reviews how changes in monsoon rainfall and temperature have affected kharif crop yields in Maharashtra over the last two decades. Rising temperatures are associated with lower yields of rice and soybean <span class="cite" contenteditable="false" data-src="sr1"></span>. Adaptation through improved varieties and water management is discussed.</p>
<p><i>Keywords:</i> climate variability, monsoon, kharif, adaptation</p>
<h1>Introduction</h1>
<p style="text-align:justify">Agriculture employs nearly half of Maharashtra’s workforce, and most kharif crops depend on the south-west monsoon. Long-term studies by the ICAR network show that each 1 °C rise in temperature can reduce yields of several crops <span class="cite" contenteditable="false" data-src="sr2"></span>. Recent years have seen record temperatures across India <span class="cite" contenteditable="false" data-src="sr3"></span>.</p>
<h1>Method</h1>
<p style="text-align:justify">District-level yield data (2003–2023) were compared with seasonal rainfall and mean temperature. Table 1 lists the districts studied.</p>
<p class="caption" data-cap="Table">Table <span class="seq" data-seq="Table" contenteditable="false">1</span>: Districts included in the study</p>
<table class="tbl t-light"><colgroup><col style="width:40%"><col style="width:30%"><col style="width:30%"></colgroup><tbody>
<tr><td><p>District</p></td><td><p>Main crop</p></td><td><p>Rainfall (mm)</p></td></tr>
<tr><td><p>Nashik</p></td><td><p>Soybean</p></td><td><p>1,034</p></td></tr>
<tr><td><p>Ratnagiri</p></td><td><p>Rice</p></td><td><p>3,365</p></td></tr>
<tr><td><p>Latur</p></td><td><p>Soybean</p></td><td><p>723</p></td></tr></tbody></table>
<h1>Results</h1>
<p style="text-align:justify">Write your findings here. Insert charts with Insert ▸ Picture and number them with References ▸ Insert caption.</p>
<h1>Discussion</h1>
<p style="text-align:justify">Interpret the results and compare them with earlier studies. Use References ▸ Insert citation to cite your sources — the reference list below updates itself, and you can switch between APA, MLA, IEEE, Harvard and Chicago at any time.</p>
<div class="biblio" contenteditable="false"></div>`, { hdr: '<p style="text-align:right"><span style="font-size:9pt">CLIMATE VARIABILITY AND KHARIF YIELDS</span></p>', ftr: '<p style="text-align:center"><span style="font-size:9pt"><span class="fld" data-fld="page" contenteditable="false">1</span></span></p>' });
        D.updateCitations(); D.updateSeq();
      },
    },
    {
      name: 'Certificate (mail merge)',
      thumbHTML: `<div class="tthumb" style="padding:7px"><div style="border:2px solid #7A4E2D;outline:1px solid #7A4E2D;outline-offset:2px;height:100%;text-align:center;padding-top:16px"><div style="font-family:'Playfair Display',serif;font-size:9px;color:#3E2716">Certificate</div><div style="font-size:4px;color:#7A4E2D;letter-spacing:1px;margin-bottom:8px">OF ACHIEVEMENT</div><div style="font-family:'Playfair Display',serif;font-size:7px;color:#7A4E2D;background:#f3ead8;display:inline-block;padding:0 3px">«Name»</div>${lines(2)}</div></div>`,
      apply() {
        D.page = Object.assign(D.defaultPage(), { orient: 'landscape', m: { t: 0.8, b: 0.8, l: 1, r: 1 } }); D.design = Object.assign(D.defaultDesign(), { styleSet: 'elegant', accent: '#7A4E2D', dark: '#3E2716', border: { style: 'art', color: '#7A4E2D', width: 3 } });
        BTS.setProps({ title: 'Certificate of achievement', subject: 'Certificate', category: 'Certificate' });
        const f = (n) => `<span class="mf" contenteditable="false" data-field="${n}">«${n}»</span>`;
        D.setContent(`
<p class="title" style="text-align:center;margin-top:36pt">Certificate</p>
<p class="subtitle" style="text-align:center;letter-spacing:4pt">OF ACHIEVEMENT</p>
<p style="text-align:center;margin-top:18pt">This certificate is proudly presented to</p>
<p style="text-align:center;margin:10pt 0"><span style="font-family:'Playfair Display';font-size:34pt;color:#7A4E2D">${f('Name')}</span></p>
<p style="text-align:center">for successfully completing <b>${f('Course')}</b> with grade <b>${f('Grade')}</b></p>
<p style="text-align:center">held at Your College, Mumbai, in ${new Date().toLocaleDateString('en-IN', { month: 'long', year: 'numeric' })}.</p>
<table class="tbl no-borders" style="width:80%;margin:44pt auto 0"><colgroup><col style="width:50%"><col style="width:50%"></colgroup><tbody>
<tr><td><p style="text-align:center">______________________</p><p style="text-align:center">Course Coordinator</p></td><td><p style="text-align:center">______________________</p><p style="text-align:center">Principal</p></td></tr></tbody></table>
<p style="text-align:center;margin-top:16pt"><span style="font-size:9pt;color:#7a8499">Go to Mailings ▸ Select recipients to load a CSV or Excel list, then Preview results and Finish &amp; merge.</span></p>`);
      },
    },
    {
      name: 'Meeting minutes',
      thumbHTML: `<div class="tthumb" style="font-family:'Source Sans 3',sans-serif"><div class="t1" style="font-weight:700;color:#1B365D">Meeting minutes</div><div style="font-size:5px;color:#2B579A;margin-bottom:5px">Student council · ${new Date().getFullYear()}</div><div style="border-top:4px solid #2B579A;margin:3px 0">${lines(3)}</div><div class="t2" style="color:#1B365D;font-weight:700">Decisions</div>${lines(3)}<div class="t2" style="color:#1B365D;font-weight:700">Action items</div><div style="border-top:4px solid #2B579A">${lines(3)}</div></div>`,
      apply() {
        D.page = D.defaultPage(); D.design = Object.assign(D.defaultDesign(), { styleSet: 'report', accent: '#2B579A', dark: '#1B365D' });
        BTS.setProps({ title: 'Meeting minutes', subject: 'Minutes', category: 'Minutes' });
        D.setContent(`
<p class="title">Meeting Minutes</p>
<p class="subtitle">Student Council · Monthly meeting</p>
<table class="tbl t-minimal t-hrow"><colgroup><col style="width:25%"><col style="width:75%"></colgroup><tbody>
<tr><td><p>Date and time</p></td><td><p>${today()}, 4:00 pm – 5:15 pm</p></td></tr>
<tr><td><p>Venue</p></td><td><p>Seminar Hall 2</p></td></tr>
<tr><td><p>Chair</p></td><td><p>Aditi Kulkarni (General Secretary)</p></td></tr>
<tr><td><p>Minutes by</p></td><td><p>Your Name</p></td></tr>
<tr><td><p>Present</p></td><td><p>Aditi K., Rohan D., Sana S., Vikram P., Neha J., Faculty advisor Prof. Desai</p></td></tr>
<tr><td><p>Absent</p></td><td><p>Kunal M. (apologies sent)</p></td></tr></tbody></table>
<h1>1. Approval of previous minutes</h1>
<p>The minutes of the meeting held last month were read and approved without changes.</p>
<h1>2. Annual tech fest</h1>
<p>Rohan presented the budget for the tech fest. Sponsorship of ₹2,50,000 has been confirmed.</p>
<ul><li>Dates fixed: second week of January.</li><li>Events to include a hackathon, robotics and a Marathi poetry slam.</li></ul>
<h1>3. Library hours during exams</h1>
<p>The council will request the librarian to keep the reading room open until 10 pm during exams.</p>
<h1>Decisions</h1>
<ol><li>Tech fest budget approved.</li><li>Letter to be sent to the librarian by Friday.</li></ol>
<h1>Action items</h1>
<table class="tbl t-header"><colgroup><col style="width:50%"><col style="width:25%"><col style="width:25%"></colgroup><tbody>
<tr><td><p>Action</p></td><td><p>Owner</p></td><td><p>Due</p></td></tr>
<tr><td><p>Send sponsorship agreements</p></td><td><p>Rohan D.</p></td><td><p>Next Monday</p></td></tr>
<tr><td><p>Draft letter to librarian</p></td><td><p>Sana S.</p></td><td><p>Friday</p></td></tr>
<tr><td><p>Book auditorium for the fest</p></td><td><p>Vikram P.</p></td><td><p>End of month</p></td></tr></tbody></table>
<p><b>Next meeting:</b> first Monday of next month, 4:00 pm, Seminar Hall 2.</p>`, { ftr: '<p style="text-align:center"><span style="font-size:9pt;color:#55627A">Page <span class="fld" data-fld="page" contenteditable="false">1</span> of <span class="fld" data-fld="pages" contenteditable="false">1</span></span></p>' });
      },
    },
  ];
  D.blankThumb = `<div class="tthumb" style="display:grid;place-items:center;color:#c3c8d2"><svg width="30" height="30" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M12 5v14M5 12h14"/></svg></div>`;
})();
