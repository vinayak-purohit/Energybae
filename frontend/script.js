// ═══════════════════════════════════════════
//  ENERGYBAE — Interactive UI Logic
// ═══════════════════════════════════════════

document.addEventListener('DOMContentLoaded', () => {
  // ── Navbar scroll effect ──
  const navbar = document.getElementById('navbar');
  window.addEventListener('scroll', () => {
    navbar.classList.toggle('scrolled', window.scrollY > 40);
  });

  // ── Smooth scroll for nav links ──
  document.querySelectorAll('a[href^="#"]').forEach(link => {
    link.addEventListener('click', e => {
      e.preventDefault();
      const target = document.querySelector(link.getAttribute('href'));
      if (target) {
        target.scrollIntoView({ behavior: 'smooth', block: 'start' });
        // Update active nav link
        document.querySelectorAll('.nav-links a').forEach(a => a.classList.remove('active'));
        const navLink = document.querySelector(`.nav-links a[href="${link.getAttribute('href')}"]`);
        if (navLink) navLink.classList.add('active');
      }
    });
  });

  // ── Active nav link on scroll ──
  const sections = document.querySelectorAll('section[id]');
  window.addEventListener('scroll', () => {
    let current = '';
    sections.forEach(section => {
      const top = section.offsetTop - 100;
      if (window.scrollY >= top) current = section.getAttribute('id');
    });
    document.querySelectorAll('.nav-links a').forEach(a => {
      a.classList.remove('active');
      if (a.getAttribute('href') === `#${current}`) a.classList.add('active');
    });
  });

  // ── Scroll animations ──
  const observer = new IntersectionObserver((entries) => {
    entries.forEach((entry, i) => {
      if (entry.isIntersecting) {
        setTimeout(() => entry.target.classList.add('visible'), i * 80);
        observer.unobserve(entry.target);
      }
    });
  }, { threshold: 0.1, rootMargin: '0px 0px -50px 0px' });
  document.querySelectorAll('.animate-on-scroll').forEach(el => observer.observe(el));

  // ── File Upload Logic ──
  const uploadZone = document.getElementById('uploadZone');
  const fileInput = document.getElementById('fileInput');
  const fileList = document.getElementById('fileList');
  const generateBtn = document.getElementById('generateBtn');
  const resultsPanel = document.getElementById('resultsPanel');
  const metricsGrid = document.getElementById('metricsGrid');

  let uploadedFiles = [];

  // Drag & drop
  uploadZone.addEventListener('dragover', e => { e.preventDefault(); uploadZone.classList.add('dragover'); });
  uploadZone.addEventListener('dragleave', () => uploadZone.classList.remove('dragover'));
  uploadZone.addEventListener('drop', e => {
    e.preventDefault();
    uploadZone.classList.remove('dragover');
    handleFiles(e.dataTransfer.files);
  });

  fileInput.addEventListener('change', () => handleFiles(fileInput.files));

  function handleFiles(files) {
    for (const file of files) {
      if (uploadedFiles.length >= 5) break;
      const allowed = ['application/pdf', 'image/png', 'image/jpeg', 'image/jpg'];
      if (!allowed.includes(file.type)) continue;
      uploadedFiles.push(file);
    }
    renderFileList();
    generateBtn.disabled = uploadedFiles.length === 0;
  }

  function renderFileList() {
    fileList.innerHTML = '';
    uploadedFiles.forEach((file, i) => {
      const sizeKB = (file.size / 1024).toFixed(1);
      const ext = file.name.split('.').pop().toUpperCase();
      const item = document.createElement('div');
      item.className = 'file-item';
      item.style.animation = `slideIn 0.3s ease ${i * 0.1}s both`;
      item.innerHTML = `
        <div class="file-name">
          <span>${ext === 'PDF' ? '📄' : '🖼️'}</span>
          <span>${file.name}</span>
          <span style="color:rgba(255,255,255,0.4)">(${sizeKB} KB)</span>
        </div>
        <button class="remove-btn" data-index="${i}" title="Remove">✕</button>
      `;
      fileList.appendChild(item);
    });

    // Remove file buttons
    fileList.querySelectorAll('.remove-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        uploadedFiles.splice(parseInt(btn.dataset.index), 1);
        renderFileList();
        generateBtn.disabled = uploadedFiles.length === 0;
      });
    });
  }

  // ── Generate Proposal ──
  let generatedExcelBlob = null;
  let generatedFileName = '';

  generateBtn.addEventListener('click', async () => {
    if (uploadedFiles.length === 0) return;

    generateBtn.disabled = true;
    generateBtn.innerHTML = '<span class="spinner"></span> AI is analyzing your bill...';
    resultsPanel.classList.remove('show');

    try {
      const formData = new FormData();
      uploadedFiles.forEach(file => formData.append('files', file));

      // For deployment: Update the string below with your actual Render URL
      const API_BASE_URL = window.location.hostname === '127.0.0.1' || window.location.hostname === 'localhost' 
        ? 'http://127.0.0.1:8000'
        : 'https://energybae-backend-rt36.onrender.com'; // <--- Live Render backend

      const response = await fetch(`${API_BASE_URL}/api/generate-proposal`, {
        method: 'POST',
        body: formData
      });

      if (!response.ok) {
        let errMsg = `Server error: ${response.status}`;
        try {
          const errData = await response.json();
          if (errData.detail) errMsg += ` - ${errData.detail}`;
        } catch(e) {}
        throw new Error(errMsg);
      }

      generatedExcelBlob = await response.blob();
      
      const contentDisposition = response.headers.get('Content-Disposition');
      if (contentDisposition) {
        const match = contentDisposition.match(/filename="(.+)"/);
        if (match) generatedFileName = match[1];
      }
      if (!generatedFileName) generatedFileName = 'Energybae_Solar_Proposal.xlsx';

      const headerData = response.headers.get('X-Extracted-Data');
      let extData = {};
      if (headerData) {
        try {
          extData = JSON.parse(headerData);
        } catch(e) {}
      }

      metricsGrid.innerHTML = `
        <div class="metric-card"><div class="label">Consumer Name</div><div class="value">${extData.consumer_name || 'N/A'}</div></div>
        <div class="metric-card"><div class="label">Consumer Number</div><div class="value">${extData.consumer_number || 'N/A'}</div></div>
        <div class="metric-card"><div class="label">Sanctioned Load</div><div class="value">${extData.sanctioned_load || 'N/A'}</div></div>
        <div class="metric-card"><div class="label">Connection Type</div><div class="value">${extData.connection_type || 'N/A'}</div></div>
        <div class="metric-card"><div class="label">Fixed Charges</div><div class="value">₹${extData.fixed_charges || 'N/A'}</div></div>
      `;

      resultsPanel.classList.add('show');
      resultsPanel.scrollIntoView({ behavior: 'smooth', block: 'center' });
    } catch (err) {
      alert('Failed to process the bill. Make sure the backend server is running.\nError: ' + err.message);
    } finally {
      generateBtn.innerHTML = '⚡ Generate Solar Proposal';
      generateBtn.disabled = false;
    }
  });

  // ── Download button ──
  document.getElementById('downloadBtn').addEventListener('click', () => {
    if (!generatedExcelBlob) {
      alert('No proposal available. Please generate one first.');
      return;
    }
    const url = window.URL.createObjectURL(generatedExcelBlob);
    const a = document.createElement('a');
    a.style.display = 'none';
    a.href = url;
    a.download = generatedFileName;
    document.body.appendChild(a);
    a.click();
    window.URL.revokeObjectURL(url);
  });

  // ── Contact form ──
  document.getElementById('contactForm').addEventListener('submit', e => {
    e.preventDefault();
    const btn = e.target.querySelector('button[type="submit"]');
    btn.textContent = '✅ Message Sent!';
    btn.style.background = 'var(--green-dark)';
    setTimeout(() => {
      btn.textContent = 'Submit';
      btn.style.background = '';
      e.target.reset();
    }, 3000);
  });

  // ── Mobile menu toggle ──
  const mobileMenu = document.getElementById('mobileMenu');
  const navLinks = document.getElementById('navLinks');
  if (mobileMenu) {
    mobileMenu.addEventListener('click', () => {
      navLinks.style.display = navLinks.style.display === 'flex' ? 'none' : 'flex';
      navLinks.style.flexDirection = 'column';
      navLinks.style.position = 'absolute';
      navLinks.style.top = '72px';
      navLinks.style.left = '0';
      navLinks.style.right = '0';
      navLinks.style.background = 'white';
      navLinks.style.padding = '20px';
      navLinks.style.boxShadow = 'var(--shadow-md)';
    });
  }

  // ── Counter animation for stats ──
  function animateCounters() {
    document.querySelectorAll('.stat-num').forEach(el => {
      const text = el.textContent;
      const match = text.match(/(\d+)/);
      if (!match) return;
      const target = parseInt(match[1]);
      const suffix = text.replace(match[1], '');
      let current = 0;
      const step = Math.max(1, Math.floor(target / 40));
      const timer = setInterval(() => {
        current += step;
        if (current >= target) { current = target; clearInterval(timer); }
        el.textContent = current + suffix;
      }, 30);
    });
  }

  // Trigger counter animation when hero is visible
  const heroObserver = new IntersectionObserver(entries => {
    if (entries[0].isIntersecting) {
      animateCounters();
      heroObserver.disconnect();
    }
  });
  const heroStats = document.querySelector('.hero-stats');
  if (heroStats) heroObserver.observe(heroStats);
});
