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
  ];
  D.blankThumb = `<div class="tthumb" style="display:grid;place-items:center;color:#c3c8d2"><svg width="30" height="30" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M12 5v14M5 12h14"/></svg></div>`;
})();
