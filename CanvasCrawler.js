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
const apiDelay = 800; // ms between API requests

// --- Helper: Paginated API fetch ---
async function pagedFetch(url, objKey, processPage) {
  let results = [];
  while (url) {
    const res = await fetch(url, { credentials: 'include' });
    if (res.status === 429) {
      const retryAfter = +res.headers.get('Retry-After') || 2;
      console.warn(`Rate limit hit, waiting ${retryAfter}s...`);
      await delay(retryAfter * 1000);
      continue;
    }
    const data = await res.json();
    if (Array.isArray(data)) results = results.concat(data);
    else if (objKey && Array.isArray(data[objKey])) results = results.concat(data[objKey]);
    if (processPage) await processPage(data, results.length);
    const link = res.headers.get('Link');
    if (link && link.includes('rel="next"')) {
      url = link.split(',').find(s => s.includes('rel="next"')).split(';')[0].replace(/[<>]/g, '').trim();
    } else {
      url = null;
    }
    await delay(apiDelay);
  }
  return results;
}

// --- Fetch only active students ---
async function fetchAllStudents(courseId) {
  const domain = 'https://canvas.ox.ac.uk';
  let url = `${domain}/api/v1/courses/${courseId}/enrollments?type[]=StudentEnrollment&state[]=active&per_page=100`;
  return await pagedFetch(url);
}

// --- Fetch all classic quizzes (published only) ---
async function fetchAllPublishedQuizzes(courseId) {
  const domain = 'https://canvas.ox.ac.uk';
  let url = `${domain}/api/v1/courses/${courseId}/quizzes?per_page=50`;
  let all = await pagedFetch(url);
  return all.filter(q => q.published);
}

// --- Fetch all quiz submissions (all pages) ---
async function fetchAllQuizSubmissions(courseId, quizId) {
  const domain = 'https://canvas.ox.ac.uk';
  let url = `${domain}/api/v1/courses/${courseId}/quizzes/${quizId}/submissions?per_page=100`;
  return await pagedFetch(url, 'quiz_submissions');
}

// --- Fetch all assignments (for New Quizzes, published only) ---
async function fetchAllAssignments(courseId) {
  const domain = 'https://canvas.ox.ac.uk';
  let url = `${domain}/api/v1/courses/${courseId}/assignments?per_page=100`;
  return await pagedFetch(url);
}

// --- Fetch all assignment submissions (paginated) ---
async function fetchAllAssignmentSubmissions(courseId, assignmentId) {
  const domain = 'https://canvas.ox.ac.uk';
  let url = `${domain}/api/v1/courses/${courseId}/assignments/${assignmentId}/submissions?per_page=100`;
  return await pagedFetch(url);
}

// --- Fetch all published discussions ---
async function fetchAllDiscussions(courseId) {
  const domain = 'https://canvas.ox.ac.uk';
  let url = `${domain}/api/v1/courses/${courseId}/discussion_topics?per_page=100`;
  let all = await pagedFetch(url);
  return all.filter(d => d.published && !d.locked);
}

// --- Fetch all discussion entries (paginated) ---
async function fetchAllDiscussionEntries(courseId, topicId) {
  const domain = 'https://canvas.ox.ac.uk';
  let url = `${domain}/api/v1/courses/${courseId}/discussion_topics/${topicId}/entries?per_page=100`;
  return await pagedFetch(url);
}

// --- Fetch all modules and module items ---
async function fetchAllModules(courseId) {
  const domain = 'https://canvas.ox.ac.uk';
  let url = `${domain}/api/v1/courses/${courseId}/modules?per_page=50`;
  return await pagedFetch(url);
}
async function fetchAllModuleItems(courseId, moduleId) {
  const domain = 'https://canvas.ox.ac.uk';
  let url = `${domain}/api/v1/courses/${courseId}/modules/${moduleId}/items?per_page=50`;
  return await pagedFetch(url);
}

// --- Fetch a page's HTML ---
async function fetchPageBody(courseId, pageUrl) {
  const domain = 'https://canvas.ox.ac.uk';
  const url = `${domain}/api/v1/courses/${courseId}/pages/${pageUrl}`;
  for (let tries = 0; tries < 3; tries++) {
    const res = await fetch(url, { credentials: 'include' });
    if (res.status === 429) {
      const retryAfter = +res.headers.get('Retry-After') || 2;
      await delay(retryAfter * 1000);
      continue;
    }
    if (!res.ok) return null;
    const data = await res.json();
    return data.body || '';
  }
  return '';
}

// === MAIN AUDIT SCRIPT ===

(async () => {
  const courseIds = [
    262596
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

      // --- Students ---
      const students = await fetchAllStudents(courseId);
      const studentIds = Array.from(new Set(students.map(s => s.user_id)));
      console.log(`[${courseName}] Found ${studentIds.length} unique active students`);

      // --- Classic Quizzes (Published Only) ---
      const quizzes = await fetchAllPublishedQuizzes(courseId);
      console.log(`[${courseName}] Found ${quizzes.length} published classic quizzes`);
      let totalStudentQuizSubs = 0;
      for (let q = 0; q < quizzes.length; q++) {
        const quiz = quizzes[q];
        console.log(`[${courseName}] Fetching submissions for classic quiz "${quiz.title}" (${q + 1} of ${quizzes.length})`);
        const submissions = await fetchAllQuizSubmissions(courseId, quiz.id);
        const subStudentIds = submissions.map(s => s.user_id).filter(id => studentIds.includes(id));
        totalStudentQuizSubs += subStudentIds.length;
      }
      let pctAnsweredQuizzes = (quizzes.length && studentIds.length)
        ? ((totalStudentQuizSubs / (studentIds.length * quizzes.length)) * 100).toFixed(1)
        : "N/A";

      // --- New Quizzes (Assignments, Published Only) ---
      // Define the same LTI domain(s) for New Quizzes detection:
      const newQuizLtiDomains = [
        'quiz-lti-dub-prod.instructure.com'
      ];
      const assignments = await fetchAllAssignments(courseId);
      const publishedNewQuizzes = assignments.filter(a =>
        a.published &&
        (a.submission_types || []).includes('external_tool') &&
        a.external_tool_url &&
        newQuizLtiDomains.some(domain => a.external_tool_url.includes(domain))
      );
      console.log(`[${courseName}] Found ${publishedNewQuizzes.length} published New Quizzes`);
      let totalStudentNewQuizSubs = 0;
      for (let nq = 0; nq < publishedNewQuizzes.length; nq++) {
        const nqAssign = publishedNewQuizzes[nq];
        console.log(`[${courseName}] Fetching submissions for New Quiz "${nqAssign.name}" (${nq + 1} of ${publishedNewQuizzes.length})`);
        const subs = await fetchAllAssignmentSubmissions(courseId, nqAssign.id);
        const subStudentIds = subs.map(s => s.user_id).filter(id => studentIds.includes(id));
        totalStudentNewQuizSubs += subStudentIds.length;
      }
      let pctAnsweredNewQuizzes = (publishedNewQuizzes.length && studentIds.length)
        ? ((totalStudentNewQuizSubs / (studentIds.length * publishedNewQuizzes.length)) * 100).toFixed(1)
        : "N/A";

      // Other published assignments = published, not new quizzes, not classic quizzes
      const otherAssignments = assignments.filter(a =>
        a.published &&
        // Not New Quizzes
        (!(a.submission_types || []).includes('external_tool') ||
        !a.external_tool_url ||
        !newQuizLtiDomains.some(domain => a.external_tool_url.includes(domain)))
        // Optionally, you can also filter out assignments linked to classic quizzes by checking assignment.quiz_id, if you want, but not strictly necessary
      );

      // % students submitting any "other" assignment (across all such assignments)
      let totalStudentOtherAssignSubs = 0;
      for (let i = 0; i < otherAssignments.length; i++) {
        const assign = otherAssignments[i];
        console.log(`[${courseName}] Fetching submissions for assignment "${assign.name}" (${i + 1} of ${otherAssignments.length})`);
        const subs = await fetchAllAssignmentSubmissions(courseId, assign.id);
        const subStudentIds = new Set(subs.map(s => s.user_id).filter(id => studentIds.includes(id)));
        totalStudentOtherAssignSubs += subStudentIds.size;
      }
      let pctAnsweredOtherAssignments = (otherAssignments.length && studentIds.length)
        ? ((totalStudentOtherAssignSubs / (studentIds.length * otherAssignments.length)) * 100).toFixed(1)
        : "N/A";

      // --- Discussions (Published Only) ---
      const discussions = await fetchAllDiscussions(courseId);
      console.log(`[${courseName}] Found ${discussions.length} published discussions`);
      let totalStudentDiscussionSubs = 0;
      for (let d = 0; d < discussions.length; d++) {
        const topic = discussions[d];
        console.log(`[${courseName}] Fetching entries for Discussion "${topic.title}" (${d + 1} of ${discussions.length})`);
        const entries = await fetchAllDiscussionEntries(courseId, topic.id);
        // Only count entries by enrolled students
        const uniqueStudentIds = new Set(entries.map(e => e.user_id).filter(id => studentIds.includes(id)));
        totalStudentDiscussionSubs += uniqueStudentIds.size;
      }
      let pctAnsweredDiscussions = (discussions.length && studentIds.length)
        ? ((totalStudentDiscussionSubs / (studentIds.length * discussions.length)) * 100).toFixed(1)
        : "N/A";

      // --- Modules/pages: count Panopto & H5P embeds ---
      const modules = await fetchAllModules(courseId);
      let panoptoPages = 0, h5pPages = 0, pageCount = 0;
      let panoptoTotal = 0, h5pTotal = 0;
      for (const module of modules) {
        const items = await fetchAllModuleItems(courseId, module.id);
        const modulePages = items.filter(i => i.type === 'Page');
        let currentPage = 0;
        for (const item of items) {
          if (item.type === 'Page') {
            currentPage++;
            console.log(`[${courseName}] Module "${module.name}": Processing page "${item.title}" (${currentPage} of ${modulePages.length})`);
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

      // --- Results ---
      results.push({
        'Course ID': courseId,
        'Course Name': courseName,
        'Number of Students': studentIds.length,
        'Classic Quizzes': quizzes.length,
        '% Students Answering Quizzes': pctAnsweredQuizzes,
        'New Quizzes': publishedNewQuizzes.length,
        '% Students Answering New Quizzes': pctAnsweredNewQuizzes,
        'Other Assignments': otherAssignments.length,
        '% Students Submitting Other Assignments': pctAnsweredOtherAssignments,
        'Discussions': discussions.length,
        '% Students Replying in Discussions': pctAnsweredDiscussions,
        'Module Pages': pageCount,
        'Pages with Panopto': panoptoPages,
        'Panopto Embeds': panoptoTotal,
        'Pages with H5P': h5pPages,
        'H5P Embeds': h5pTotal
      });
      console.log(
        `Done: "${courseName}" - % students answering classic quizzes: ${pctAnsweredQuizzes}%, new quizzes: ${pctAnsweredNewQuizzes}%, discussions: ${pctAnsweredDiscussions}%`
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
        '% Students Answering New Quizzes': 'ERR',
        'Other Assignments': 'ERR',
        '% Students Submitting Other Assignments': 'ERR',
        'Discussions': 'ERR',
        '% Students Replying in Discussions': 'ERR',
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
