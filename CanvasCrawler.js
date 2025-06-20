// --- Helper to parse and count Panopto/H5P embeds in page HTML ---
function countEmbedsInHTML(html, type) {
  let count = 0;
  const parser = new DOMParser();
  const doc = parser.parseFromString(html, 'text/html');
  const iframes = doc.querySelectorAll('iframe');
  iframes.forEach(iframe => {
    const src = iframe.getAttribute('src') || '';
    if (type === "panopto") {
      if (src.toLowerCase().includes('panopto') || (
        src.includes('/external_tools/retrieve') &&
        decodeURIComponent((src.split('url=')[1]||'').split('&')[0]||'').toLowerCase().includes('panopto')
      )) count++;
    }
    if (type === "h5p") {
      if (src.toLowerCase().includes('h5p') || (
        src.includes('/external_tools/retrieve') &&
        decodeURIComponent((src.split('url=')[1]||'').split('&')[0]||'').toLowerCase().includes('h5p')
      )) count++;
    }
  });
  const links = doc.querySelectorAll('a');
  links.forEach(a => {
    const href = a.getAttribute('href') || '';
    if (type === "panopto" && href.toLowerCase().includes('panopto')) count++;
    if (type === "h5p" && href.toLowerCase().includes('h5p')) count++;
  });
  return count;
}

// --- Helper: Delay ---
const delay = ms => new Promise(res => setTimeout(res, ms));
const apiDelay = 800; // ms between API requests, increase if you get rate limits

// --- Helper: Fetch all pages of enrollments (students only) ---
async function fetchAllStudents(courseId) {
  const domain = 'https://canvas.ox.ac.uk';
  let students = [];
  let url = `${domain}/api/v1/courses/${courseId}/enrollments?type[]=StudentEnrollment&per_page=100`;
  while (url) {
    const res = await fetch(url, { credentials: 'include' });
    if (res.status === 429) {
      const retryAfter = +res.headers.get('Retry-After') || 2;
      console.warn(`Rate limit on enrollments, waiting ${retryAfter}s...`);
      await delay(retryAfter * 1000);
      continue;
    }
    const data = await res.json();
    students = students.concat(data);
    // Pagination
    const link = res.headers.get('Link');
    if (link && link.includes('rel="next"')) {
      url = link.split(',').find(s => s.includes('rel="next"')).split(';')[0].replace(/[<>]/g, '').trim();
    } else {
      url = null;
    }
    await delay(apiDelay);
  }
  return students;
}

// --- Helper: Fetch all classic quizzes ---
async function fetchAllQuizzes(courseId) {
  const domain = 'https://canvas.ox.ac.uk';
  let quizzes = [];
  let url = `${domain}/api/v1/courses/${courseId}/quizzes?per_page=50`;
  while (url) {
    const res = await fetch(url, { credentials: 'include' });
    if (res.status === 429) {
      const retryAfter = +res.headers.get('Retry-After') || 2;
      console.warn(`Rate limit on quizzes, waiting ${retryAfter}s...`);
      await delay(retryAfter * 1000);
      continue;
    }
    const data = await res.json();
    quizzes = quizzes.concat(data);
    // Pagination
    const link = res.headers.get('Link');
    if (link && link.includes('rel="next"')) {
      url = link.split(',').find(s => s.includes('rel="next"')).split(';')[0].replace(/[<>]/g, '').trim();
    } else {
      url = null;
    }
    await delay(apiDelay);
  }
  return quizzes;
}

// --- Helper: Fetch all quiz submissions (paginated) ---
async function fetchAllQuizSubmissions(courseId, quizId) {
  const domain = 'https://canvas.ox.ac.uk';
  let submissions = [];
  let url = `${domain}/api/v1/courses/${courseId}/quizzes/${quizId}/submissions?per_page=100`;
  while (url) {
    const res = await fetch(url, { credentials: 'include' });
    if (res.status === 429) {
      const retryAfter = +res.headers.get('Retry-After') || 2;
      console.warn(`Rate limit on quiz submissions, waiting ${retryAfter}s...`);
      await delay(retryAfter * 1000);
      continue;
    }
    const data = await res.json();
    submissions = submissions.concat(data.quiz_submissions || []);
    // Pagination
    const link = res.headers.get('Link');
    if (link && link.includes('rel="next"')) {
      url = link.split(',').find(s => s.includes('rel="next"')).split(';')[0].replace(/[<>]/g, '').trim();
    } else {
      url = null;
    }
    await delay(apiDelay);
  }
  return submissions;
}

// --- Helper: Fetch all assignments (including New Quizzes) ---
async function fetchAllAssignments(courseId) {
  const domain = 'https://canvas.ox.ac.uk';
  let assignments = [];
  let url = `${domain}/api/v1/courses/${courseId}/assignments?per_page=100`;
  while (url) {
    const res = await fetch(url, { credentials: 'include' });
    if (res.status === 429) {
      const retryAfter = +res.headers.get('Retry-After') || 2;
      console.warn(`Rate limit on assignments, waiting ${retryAfter}s...`);
      await delay(retryAfter * 1000);
      continue;
    }
    const data = await res.json();
    assignments = assignments.concat(data);
    // Pagination
    const link = res.headers.get('Link');
    if (link && link.includes('rel="next"')) {
      url = link.split(',').find(s => s.includes('rel="next"')).split(';')[0].replace(/[<>]/g, '').trim();
    } else {
      url = null;
    }
    await delay(apiDelay);
  }
  return assignments;
}

// --- Helper: Fetch all modules and module items ---
async function fetchAllModules(courseId) {
  const domain = 'https://canvas.ox.ac.uk';
  let modules = [];
  let url = `${domain}/api/v1/courses/${courseId}/modules?per_page=50`;
  while (url) {
    const res = await fetch(url, { credentials: 'include' });
    if (res.status === 429) {
      const retryAfter = +res.headers.get('Retry-After') || 2;
      console.warn(`Rate limit on modules, waiting ${retryAfter}s...`);
      await delay(retryAfter * 1000);
      continue;
    }
    const data = await res.json();
    modules = modules.concat(data);
    // Pagination
    const link = res.headers.get('Link');
    if (link && link.includes('rel="next"')) {
      url = link.split(',').find(s => s.includes('rel="next"')).split(';')[0].replace(/[<>]/g, '').trim();
    } else {
      url = null;
    }
    await delay(apiDelay);
  }
  return modules;
}

async function fetchAllModuleItems(courseId, moduleId) {
  const domain = 'https://canvas.ox.ac.uk';
  let items = [];
  let url = `${domain}/api/v1/courses/${courseId}/modules/${moduleId}/items?per_page=50`;
  while (url) {
    const res = await fetch(url, { credentials: 'include' });
    if (res.status === 429) {
      const retryAfter = +res.headers.get('Retry-After') || 2;
      console.warn(`Rate limit on module items, waiting ${retryAfter}s...`);
      await delay(retryAfter * 1000);
      continue;
    }
    const data = await res.json();
    items = items.concat(data);
    // Pagination
    const link = res.headers.get('Link');
    if (link && link.includes('rel="next"')) {
      url = link.split(',').find(s => s.includes('rel="next"')).split(';')[0].replace(/[<>]/g, '').trim();
    } else {
      url = null;
    }
    await delay(apiDelay);
  }
  return items;
}

// --- Helper: Fetch a page's HTML ---
async function fetchPageBody(courseId, pageUrl) {
  const domain = 'https://canvas.ox.ac.uk';
  const url = `${domain}/api/v1/courses/${courseId}/pages/${pageUrl}`;
  for (let tries = 0; tries < 3; tries++) {
    const res = await fetch(url, { credentials: 'include' });
    if (res.status === 429) {
      const retryAfter = +res.headers.get('Retry-After') || 2;
      console.warn(`Rate limit on page body, waiting ${retryAfter}s...`);
      await delay(retryAfter * 1000);
      continue;
    }
    if (!res.ok) return null;
    const data = await res.json();
    return data.body || '';
  }
  return '';
}

// --- Main script ---
(async () => {
  const courseIds = [
    262594
  ];
  const domain = 'https://canvas.ox.ac.uk';

  const results = [];

  for (let c = 0; c < courseIds.length; c++) {
    const courseId = courseIds[c];
    console.log(`\n=== [${c + 1}/${courseIds.length}] Auditing course ${courseId} ===`);
    try {
      // Get course details
      const detailsRes = await fetch(`${domain}/api/v1/courses/${courseId}`, { credentials: 'include' });
      const details = detailsRes.ok ? await detailsRes.json() : {};
      const courseName = details.name || 'Unknown';

      // Get students
      const students = await fetchAllStudents(courseId);
      const studentIds = students.map(s => s.user_id);
      console.log(`Course "${courseName}": ${studentIds.length} students`);

      // Get quizzes
      const quizzes = await fetchAllQuizzes(courseId);
      console.log(`Found ${quizzes.length} classic quizzes`);

      // For % of students answering quizzes
      let totalStudentQuizSubs = 0;
      let quizSubMatrix = {}; // quiz_id -> Set of student_ids

      for (let q = 0; q < quizzes.length; q++) {
        const quiz = quizzes[q];
        const submissions = await fetchAllQuizSubmissions(courseId, quiz.id);
        console.log(`Quiz "${quiz.title}": ${submissions.length} submissions`);
        // Only count submissions by enrolled students (should always be true)
        const subStudentIds = submissions.map(s => s.user_id).filter(id => studentIds.includes(id));
        quizSubMatrix[quiz.id] = new Set(subStudentIds);
        totalStudentQuizSubs += subStudentIds.length;
      }

      let pctAnsweredQuizzes = (quizzes.length && studentIds.length)
        ? ((totalStudentQuizSubs / (studentIds.length * quizzes.length)) * 100).toFixed(1)
        : "N/A";

      // Get assignments (New Quizzes and other assignments)
      const assignments = await fetchAllAssignments(courseId);
      const newQuizzes = assignments.filter(a =>
        (a.submission_types || []).includes('external_tool') &&
        a.external_tool_url && a.external_tool_url.includes('new_quizzes')
      );
      let newQuizzesWithSubs = 0;
      for (const nq of newQuizzes) {
        const subsRes = await fetch(`${domain}/api/v1/courses/${courseId}/assignments/${nq.id}/submissions?per_page=100`, { credentials: 'include' });
        let subs = subsRes.ok ? await subsRes.json() : [];
        if (Array.isArray(subs) && subs.length > 0) newQuizzesWithSubs++;
        await delay(apiDelay);
      }
      const otherAssigns = assignments.filter(a =>
        !(a.submission_types || []).includes('external_tool') ||
        !a.external_tool_url || !a.external_tool_url.includes('new_quizzes')
      );
      let otherAssignsWithSubs = 0;
      for (const a of otherAssigns) {
        const subsRes = await fetch(`${domain}/api/v1/courses/${courseId}/assignments/${a.id}/submissions?per_page=100`, { credentials: 'include' });
        let subs = subsRes.ok ? await subsRes.json() : [];
        if (Array.isArray(subs) && subs.length > 0) otherAssignsWithSubs++;
        await delay(apiDelay);
      }

      // Modules/pages: count Panopto & H5P embeds
      const modules = await fetchAllModules(courseId);
      let panoptoPages = 0, h5pPages = 0, pageCount = 0;
      let panoptoTotal = 0, h5pTotal = 0;
      for (const module of modules) {
        const items = await fetchAllModuleItems(courseId, module.id);
        for (const item of items) {
          if (item.type === 'Page') {
            pageCount++;
            const pageBody = await fetchPageBody(courseId, item.page_url);
            if (pageBody) {
              const panoptoOnPage = countEmbedsInHTML(pageBody, "panopto");
              const h5pOnPage = countEmbedsInHTML(pageBody, "h5p");
              if (panoptoOnPage) panoptoPages++;
              if (h5pOnPage) h5pPages++;
              panoptoTotal += panoptoOnPage;
              h5pTotal += h5pOnPage;
            }
            await delay(apiDelay);
          }
        }
      }

      // Record for table
      results.push({
        'Course ID': courseId,
        'Course Name': courseName,
        'Number of Students': studentIds.length,
        'Classic Quizzes': quizzes.length,
        '% Students Answering Quizzes': pctAnsweredQuizzes,
        'New Quizzes': newQuizzes.length,
        'New Quizzes w/ Subs': newQuizzesWithSubs,
        'Other Assignments': otherAssigns.length,
        'Other Assignments w/ Subs': otherAssignsWithSubs,
        'Module Pages': pageCount,
        'Pages with Panopto': panoptoPages,
        'Panopto Embeds': panoptoTotal,
        'Pages with H5P': h5pPages,
        'H5P Embeds': h5pTotal
      });
      console.log(
        `Done: "${courseName}": % students answering quizzes = ${pctAnsweredQuizzes}%`
      );
    } catch (err) {
      console.error(`Error processing course ${courseId}:`, err);
      results.push({
        'Course ID': courseId,
        'Course Name': 'ERROR',
        'Number of Students': 'ERR',
        'Classic Quizzes': 'ERR',
        '% Students Answering Quizzes': 'ERR',
        'New Quizzes': 'ERR',
        'New Quizzes w/ Subs': 'ERR',
        'Other Assignments': 'ERR',
        'Other Assignments w/ Subs': 'ERR',
        'Module Pages': 'ERR',
        'Pages with Panopto': 'ERR',
        'Panopto Embeds': 'ERR',
        'Pages with H5P': 'ERR',
        'H5P Embeds': 'ERR'
      });
    }
  }

  console.table(results);

  // Optional: CSV download
  const csv = [
    Object.keys(results[0]).join(','),
    ...results.map(r => Object.values(r).join(','))
  ].join('\n');

  const blob = new Blob([csv], { type: 'text/csv' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'canvas_course_audit.csv';
  a.click();
})();
