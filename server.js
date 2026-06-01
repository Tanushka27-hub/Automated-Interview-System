const express = require('express');
const twilio = require('twilio');
const bodyParser = require('body-parser');
const cors = require('cors');
const path = require('path');
const multer = require('multer');
const XLSX = require('xlsx');
const Papa = require('papaparse');
const fs = require('fs');
const schedule = require('node-schedule');
const moment = require('moment-timezone');

const NGROK_URL = 'https://agreeably-garland-jaybird.ngrok-free.dev'; // Updated ngrok URL

const app = express();
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));
app.use(cors());

app.use(express.static(path.join(__dirname, 'public')));

// Configure multer for file uploads
const upload = multer({ dest: 'uploads/' });

require('dotenv').config();

const accountSid = process.env.TWILIO_ACCOUNT_SID;
const authToken = process.env.TWILIO_AUTH_TOKEN;

// In-memory storage
let candidates = [];
let bookedTimeSlots = {};

const departments = ['HR', 'IT', 'E Department', 'DISP'];

const subDepartments = {
  HR: ['Recruitment', 'Employee Relations'],
  IT: ['Development', 'Support'],
  'E Department': ['Operations', 'Analytics'],
  DISP: ['Operations', 'PDispatch']
};

const sections = {
  HR: {
    Recruitment: ['Sourcing', 'Onboarding'],
    'Employee Relations': ['Engagement', 'Compliance']
  },
  IT: {
    Development: ['Frontend', 'Backend'],
    Support: ['Technical', 'Customer']
  },
  'E Department': {
    Operations: [],
    Analytics: []
  },
  DISP: {
    Operations: ['MM HotZone', 'MM ColdZone'],
    PDispatch: ['MM HotZone', 'MM ColdZone']
  }
};

const designations = {
  HR: {
    Recruitment: {
      Sourcing: ['Assistant Manager', 'Deputy Manager', 'Manager'],
      Onboarding: ['Assistant Manager', 'Deputy Manager', 'Manager']
    },
    'Employee Relations': {
      Engagement: ['Assistant Manager', 'Deputy Manager', 'Manager'],
      Compliance: ['Assistant Manager', 'Deputy Manager', 'Manager']
    }
  },
  IT: {
    Development: {
      Frontend: ['Junior Developer', 'Senior Developer', 'Tech Lead'],
      Backend: ['Junior Developer', 'Senior Developer', 'Tech Lead']
    },
    Support: {
      Technical: ['Support Engineer', 'Senior Support Engineer', 'Support Manager'],
      Customer: ['Support Engineer', 'Senior Support Engineer', 'Support Manager']
    }
  },
  'E Department': {
    Operations: {
      Operations: ['Coordinator', 'Specialist', 'Manager']
    },
    Analytics: {
      Analytics: ['Coordinator', 'Specialist', 'Manager']
    }
  },
  DISP: {
    Operations: {
      'MM HotZone': ['Manager', 'Senior Manager', 'AGM'],
      'MM ColdZone': ['Manager', 'Senior Manager', 'AGM']
    },
    PDispatch: {
      'MM HotZone': ['Manager', 'Senior Manager', 'AGM'],
      'MM ColdZone': ['Manager', 'Senior Manager', 'AGM']
    }
  }
};

const questionBank = {
  HR: {
    Recruitment: {
      Sourcing: {
        'Assistant Manager': [
          { question: 'What is the primary goal of sourcing? 1. Increase diversity 2. Fill vacancies quickly 3. Reduce costs', correctAnswer: '1' },
          { question: 'Which method is best for sourcing candidates? 1. Job boards 2. Referrals 3. Social media', correctAnswer: '2' }
        ],
        'Deputy Manager': [
          { question: 'What is key in candidate sourcing? 1. Resume review 2. Networking 3. Both', correctAnswer: '3' },
          { question: 'What improves sourcing efficiency? 1. Automation 2. Manual search 3. More staff', correctAnswer: '1' }
        ],
        Manager: [
          { question: 'What defines a successful sourcing strategy? 1. Quick hires 2. Quality hires 3. Low cost', correctAnswer: '2' },
          { question: 'Best way to assess sourcing channels? 1. Analytics 2. Intuition 3. Feedback', correctAnswer: '1' }
        ]
      },
      Onboarding: {
        'Assistant Manager': [
          { question: 'What is the goal of onboarding? 1. Compliance 2. Employee integration 3. Cost saving', correctAnswer: '2' },
          { question: 'What improves onboarding? 1. Training 2. Isolation 3. Strict rules', correctAnswer: '1' }
        ],
        'Deputy Manager': [
          { question: 'What is key in onboarding? 1. Documentation 2. Engagement 3. Both', correctAnswer: '3' },
          { question: 'What reduces onboarding time? 1. Automation 2. Long meetings 3. More staff', correctAnswer: '1' }
        ],
        Manager: [
          { question: 'What measures onboarding success? 1. Speed 2. Retention 3. Cost', correctAnswer: '2' },
          { question: 'Best way to improve onboarding? 1. Feedback 2. More rules 3. Less training', correctAnswer: '1' }
        ]
      }
    },
    'Employee Relations': {
      Engagement: {
        'Assistant Manager': [
          { question: 'What promotes engagement? 1. Rewards 2. Isolation 3. Strict rules', correctAnswer: '1' },
          { question: 'What is employee engagement? 1. Satisfaction 2. Productivity 3. Both', correctAnswer: '3' }
        ],
        'Deputy Manager': [
          { question: 'What improves engagement? 1. Recognition 2. Overtime 3. Micromanagement', correctAnswer: '1' },
          { question: 'What is a pulse survey? 1. Feedback tool 2. Financial audit 3. Health check', correctAnswer: '1' }
        ],
        Manager: [
          { question: 'What reduces turnover? 1. Engagement programs 2. Higher pay 3. More rules', correctAnswer: '1' },
          { question: 'Best way to measure engagement? 1. Surveys 2. Sales 3. Attendance', correctAnswer: '1' }
        ]
      },
      Compliance: {
        'Assistant Manager': [
          { question: 'What ensures compliance? 1. Policies 2. Incentives 3. Ignoring issues', correctAnswer: '1' },
          { question: 'What is a compliance audit? 1. Policy check 2. Financial review 3. Marketing plan', correctAnswer: '1' }
        ],
        'Deputy Manager': [
          { question: 'What is a labor law? 1. Employment regulation 2. Tax code 3. Marketing rule', correctAnswer: '1' },
          { question: 'What prevents violations? 1. Training 2. Bonuses 3. Overtime', correctAnswer: '1' }
        ],
        Manager: [
          { question: 'What is risk in compliance? 1. Legal issues 2. Sales loss 3. Staff shortage', correctAnswer: '1' },
          { question: 'Best way to enforce compliance? 1. Audits 2. Rewards 3. Less training', correctAnswer: '1' }
        ]
      }
    }
  },
  IT: {
    Development: {
      Frontend: {
        'Junior Developer': [
          { question: 'What is HTML? 1. Markup language 2. Programming language 3. Database', correctAnswer: '1' },
          { question: 'What is CSS used for? 1. Styling 2. Logic 3. Data storage', correctAnswer: '1' }
        ],
        'Senior Developer': [
          { question: 'What is a framework? 1. Code structure 2. Database 3. Hardware', correctAnswer: '1' },
          { question: 'What is responsive design? 1. Device adaptability 2. Code optimization 3. Security', correctAnswer: '1' }
        ],
        'Tech Lead': [
          { question: 'What is a component library? 1. Reusable UI elements 2. Database schema 3. Server code', correctAnswer: '1' },
          { question: 'What ensures frontend performance? 1. Optimization 2. More servers 3. Less code', correctAnswer: '1' }
        ]
      },
      Backend: {
        'Junior Developer': [
          { question: 'What is an API? 1. Interface for services 2. User interface 3. Hardware', correctAnswer: '1' },
          { question: 'What is a database? 1. Data storage 2. Code library 3. UI tool', correctAnswer: '1' }
        ],
        'Senior Developer': [
          { question: 'What is REST? 1. API standard 2. Database type 3. Frontend framework', correctAnswer: '1' },
          { question: 'What is scalability? 1. System growth 2. Code reduction 3. UI design', correctAnswer: '1' }
        ],
        'Tech Lead': [
          { question: 'What is microservices? 1. Modular architecture 2. Single app 3. UI pattern', correctAnswer: '1' },
          { question: 'What ensures backend reliability? 1. Redundancy 2. Less code 3. More staff', correctAnswer: '1' }
        ]
      }
    },
    Support: {
      Technical: {
        'Support Engineer': [
          { question: 'What is a ticket? 1. Support request 2. Software license 3. Hardware part', correctAnswer: '1' },
          { question: 'What is SLA? 1. Service level agreement 2. Software lifecycle 3. System load', correctAnswer: '1' }
        ],
        'Senior Support Engineer': [
          { question: 'What is root cause analysis? 1. Issue source 2. User training 3. System upgrade', correctAnswer: '1' },
          { question: 'What improves response time? 1. Automation 2. Manual checks 3. More staff', correctAnswer: '1' }
        ],
        'Support Manager': [
          { question: 'What measures support success? 1. Resolution time 2. Ticket volume 3. Complaints', correctAnswer: '1' },
          { question: 'What is escalation? 1. Issue elevation 2. User feedback 3. Software update', correctAnswer: '1' }
        ]
      },
      Customer: {
        'Support Engineer': [
          { question: 'What is customer support? 1. User assistance 2. Code writing 3. Hardware repair', correctAnswer: '1' },
          { question: 'What is a CRM? 1. Customer management tool 2. Code editor 3. Database', correctAnswer: '1' }
        ],
        'Senior Support Engineer': [
          { question: 'What improves customer satisfaction? 1. Quick response 2. Long calls 3. Less contact', correctAnswer: '1' },
          { question: 'What is a knowledge base? 1. Help resource 2. Code library 3. Financial record', correctAnswer: '1' }
        ],
        'Support Manager': [
          { question: 'What is churn rate? 1. Customer loss 2. Ticket volume 3. Staff turnover', correctAnswer: '1' },
          { question: 'What enhances support? 1. Training 2. Less staff 3. More tickets', correctAnswer: '1' }
        ]
      }
    }
  },
  'E Department': {
    Operations: {
      Coordinator: [
        { question: 'What is the primary goal of operations? 1. Efficiency 2. Marketing 3. Sales', correctAnswer: '1' },
        { question: 'What improves operational workflows? 1. Automation 2. Manual processes 3. More staff', correctAnswer: '1' }
      ],
      Specialist: [
        { question: 'What is key in operations management? 1. Process optimization 2. Customer support 3. Financial audits', correctAnswer: '1' },
        { question: 'What reduces operational costs? 1. Technology 2. Increased staff 3. Longer hours', correctAnswer: '1' }
      ],
      Manager: [
        { question: 'What defines successful operations? 1. High efficiency 2. More products 3. Lower prices', correctAnswer: '1' },
        { question: 'How to measure operational success? 1. KPIs 2. Customer feedback 3. Revenue', correctAnswer: '1' }
      ]
    },
    Analytics: {
      Coordinator: [
        { question: 'What is data analytics? 1. Data interpretation 2. Software development 3. Hardware maintenance', correctAnswer: '1' },
        { question: 'What tool is used for analytics? 1. Excel 2. Word 3. Photoshop', correctAnswer: '1' }
      ],
      Specialist: [
        { question: 'What is predictive analytics? 1. Forecasting trends 2. Code testing 3. UI design', correctAnswer: '1' },
        { question: 'What improves analytics accuracy? 1. Quality data 2. More staff 3. Faster computers', correctAnswer: '1' }
      ],
      Manager: [
        { question: 'What is a data-driven decision? 1. Based on analytics 2. Based on intuition 3. Based on surveys', correctAnswer: '1' },
        { question: 'How to ensure analytics reliability? 1. Data validation 2. More reports 3. Less data', correctAnswer: '1' }
      ]
    }
  },
  DISP: {
    Operations: {
      'MM HotZone': {
        Manager: [
          { question: 'What is the primary goal of MM HotZone operations? 1. Rapid response 2. Cost reduction 3. Marketing', correctAnswer: '1' },
          { question: 'What improves MM HotZone efficiency? 1. Real-time tracking 2. Manual logs 3. More staff', correctAnswer: '1' }
        ],
        'Senior Manager': [
          { question: 'What is key in MM HotZone management? 1. Coordination 2. Financial audits 3. Customer support', correctAnswer: '1' },
          { question: 'What reduces MM HotZone delays? 1. Automation 2. Increased staff 3. Longer hours', correctAnswer: '1' }
        ],
        AGM: [
          { question: 'What defines MM HotZone success? 1. Timely dispatch 2. More vehicles 3. Lower prices', correctAnswer: '1' },
          { question: 'How to measure MM HotZone performance? 1. KPIs 2. Customer feedback 3. Revenue', correctAnswer: '1' }
        ]
      },
      'MM ColdZone': {
        Manager: [
          { question: 'What is the primary goal of MM ColdZone operations? 1. Safe storage 2. Quick delivery 3. Marketing', correctAnswer: '1' },
          { question: 'What improves MM ColdZone efficiency? 1. Temperature control 2. Manual checks 3. More staff', correctAnswer: '1' }
        ],
        'Senior Manager': [
          { question: 'What is key in MM ColdZone management? 1. Compliance 2. Financial audits 3. Customer support', correctAnswer: '1' },
          { question: 'What ensures MM ColdZone reliability? 1. Monitoring systems 2. Increased staff 3. Longer hours', correctAnswer: '1' }
        ],
        AGM: [
          { question: 'What defines MM ColdZone success? 1. Storage integrity 2. More vehicles 3. Lower prices', correctAnswer: '1' },
          { question: 'How to measure MM ColdZone performance? 1. KPIs 2. Customer feedback 3. Revenue', correctAnswer: '1' }
        ]
      }
    },
    PDispatch: {
      'MM HotZone': {
        Manager: [
          { question: 'What is the primary goal of MM HotZone PDispatch? 1. Speedy dispatch 2. Cost reduction 3. Marketing', correctAnswer: '1' },
          { question: 'What improves MM HotZone PDispatch efficiency? 1. Route optimization 2. Manual planning 3. More staff', correctAnswer: '1' }
        ],
        'Senior Manager': [
          { question: 'What is key in MM HotZone PDispatch management? 1. Coordination 2. Financial audits 3. Customer support', correctAnswer: '1' },
          { question: 'What reduces MM HotZone PDispatch delays? 1. Automation 2. Increased staff 3. Longer hours', correctAnswer: '1' }
        ],
        AGM: [
          { question: 'What defines MM HotZone PDispatch success? 1. On-time delivery 2. More vehicles 3. Lower prices', correctAnswer: '1' },
          { question: 'How to measure MM HotZone PDispatch performance? 1. KPIs 2. Customer feedback 3. Revenue', correctAnswer: '1' }
        ]
      },
      'MM ColdZone': {
        Manager: [
          { question: 'What is the primary goal of MM ColdZone PDispatch? 1. Safe transport 2. Quick delivery 3. Marketing', correctAnswer: '1' },
          { question: 'What improves MM ColdZone PDispatch efficiency? 1. Cold chain logistics 2. Manual checks 3. More staff', correctAnswer: '1' }
        ],
        'Senior Manager': [
          { question: 'What is key in MM ColdZone PDispatch management? 1. Compliance 2. Financial audits 3. Customer support', correctAnswer: '1' },
          { question: 'What ensures MM ColdZone PDispatch reliability? 1. Monitoring systems 2. Increased staff 3. Longer hours', correctAnswer: '1' }
        ],
        AGM: [
          { question: 'What defines MM ColdZone PDispatch success? 1. Delivery integrity 2. More vehicles 3. Lower prices', correctAnswer: '1' },
          { question: 'How to measure MM ColdZone PDispatch performance? 1. KPIs 2. Customer feedback 3. Revenue', correctAnswer: '1' }
        ]
      }
    }
  }
};

// Generate time slots from 10:00 AM to 11:00 AM and 3:00 PM to 8:00 PM IST
function generateTimeSlots() {
  const timeSlots = [];
  let morningStart = moment.tz('Asia/Kolkata').set({ hour: 10, minute: 0, second: 0, millisecond: 0 });
  const morningEnd = moment.tz('Asia/Kolkata').set({ hour: 11, minute: 0, second: 0, millisecond: 0 });
  while (morningStart < morningEnd) {
    timeSlots.push(morningStart.format('h:mm A'));
    morningStart.add(10, 'minutes');
  }
  let afternoonStart = moment.tz('Asia/Kolkata').set({ hour: 15, minute: 0, second: 0, millisecond: 0 });
  const afternoonEnd = moment.tz('Asia/Kolkata').set({ hour: 20, minute: 0, second: 0, millisecond: 0 });
  while (afternoonStart <= afternoonEnd) {
    timeSlots.push(afternoonStart.format('h:mm A'));
    afternoonStart.add(15, 'minutes');
  }
  return timeSlots;
}

// Get available time slots for a given date
function getAvailableTimeSlots(date) {
  const allTimeSlots = generateTimeSlots();
  const booked = bookedTimeSlots[date] || [];
  return allTimeSlots.filter(slot => !booked.includes(slot));
}

// Promotion logic
const getNextDesignation = (dept, subDept, section, currentDesignation) => {
  if (!dept || dept === 'Not selected' || !subDept || subDept === 'Not selected' || !currentDesignation || currentDesignation === 'Not selected') {
    console.warn(`[getNextDesignation] Invalid input: dept=${dept}, subDept=${subDept}, section=${section}, currentDesignation=${currentDesignation}`);
    return currentDesignation;
  }

  const designationKey = dept === 'E Department' ? subDept : section;
  if (!designationKey || designationKey === 'Not selected') {
    console.warn(`[getNextDesignation] Invalid designationKey: ${designationKey}`);
    return currentDesignation;
  }

  const designationList = designations[dept]?.[subDept]?.[designationKey];
  if (!Array.isArray(designationList)) {
    console.warn(`[getNextDesignation] Designation list not found for dept=${dept}, subDept=${subDept}, designationKey=${designationKey}`);
    return currentDesignation;
  }

  const currentIndex = designationList.indexOf(currentDesignation);
  if (currentIndex === -1) {
    console.warn(`[getNextDesignation] Current designation ${currentDesignation} not found in list:`, designationList);
    return currentDesignation;
  }

  if (currentIndex < designationList.length - 1) {
    return designationList[currentIndex + 1];
  }
  return currentDesignation;
};

// Schedule interview at specified date and time in IST
function scheduleInterview(candidate) {
  const [hours, minutesAmpm] = candidate.interviewTime.split(':');
  const [minutes, ampm] = minutesAmpm.split(' ');
  let hour = parseInt(hours);
  if (ampm === 'PM' && hour !== 12) hour += 12;
  if (ampm === 'AM' && hour === 12) hour = 0;
  const [year, month, day] = candidate.interviewDate.split('-').map(Number);

  const scheduleTime = moment.tz(
    `${year}-${month}-${day} ${hour}:${minutes}`,
    'YYYY-M-D H:mm',
    'Asia/Kolkata'
  ).toDate();

  console.log(`[SCHEDULE-INTERVIEW] Scheduling for ${candidate.phone} at ${scheduleTime.toString()} (IST)`);

  const job = schedule.scheduleJob(scheduleTime, async () => {
    const candidateIndex = candidates.findIndex(c => c.phone === candidate.phone);
    if (candidateIndex === -1 || !candidates[candidateIndex].status.startsWith('Scheduled')) {
      console.log(`[SCHEDULE-INTERVIEW] Candidate ${candidate.phone} not found or not Scheduled`);
      return;
    }

    console.log(`[SCHEDULE-INTERVIEW] Job triggered for ${candidate.phone} at ${new Date().toString()}`);
    candidates[candidateIndex].status = 'Interview Started';
    candidates[candidateIndex].step = 'question';
    candidates[candidateIndex].currentQuestionIndex = 0;
    candidates[candidateIndex].answers = [];
    candidates[candidateIndex].correctAnswers = 0;

    try {
      await client.calls.create({
        url: `${NGROK_URL}/voice-question`,
        to: candidate.phone,
        from: twilioNumber,
        statusCallback: `${NGROK_URL}/call-status`,
        statusCallbackEvent: ['completed']
      });
      console.log(`[SCHEDULE-INTERVIEW] Voice call initiated for ${candidate.phone}`);
    } catch (error) {
      console.error(`[SCHEDULE-INTERVIEW] Error start interview for ${candidate.phone}:`, error);
      candidates[candidateIndex].status = 'Failed';
      await client.messages.create({
        body: 'An error occurred during your interview. Please contact support.',
        from: whatsappNumber,
        to: `whatsapp:${candidate.phone}`
      });
    }
  });

  console.log(`[SCHEDULE-INTERVIEW] Job scheduled for ${candidate.phone}: ${job.name}`);
}

// Upload candidates from Excel or CSV
app.post('/upload-candidates', upload.single('file'), (req, res) => {
  try {
    if (!req.file) {
      console.error('[UPLOAD-CANDIDATES] No file uploaded');
      return res.status(400).json({ error: 'No file uploaded' });
    }

    const fileExtension = req.file.originalname.split('.').pop().toLowerCase();
    let data = [];

    if (fileExtension === 'xlsx' || fileExtension === 'xls') {
      const workbook = XLSX.readFile(req.file.path);
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      data = XLSX.utils.sheet_to_json(sheet);
    } else if (fileExtension === 'csv') {
      const fileContent = fs.readFileSync(req.file.path, 'utf8');
      const parseResult = Papa.parse(fileContent, { header: true, skipEmptyLines: true });
      data = parseResult.data;
    } else {
      console.error('[UPLOAD-CANDIDATES] Unsupported file format');
      return res.status(400).json({ error: 'Unsupported file format. Use Excel (.xlsx, .xls) or CSV (.csv)' });
    }

    fs.unlinkSync(req.file.path);

    if (data.length === 0) {
      console.error('[UPLOAD-CANDIDATES] No valid candidates found');
      return res.status(400).json({ error: 'File is empty' });
    }

    const validMethods = ['voice', 'whatsapp'];
    const newCandidates = [];
    const errors = [];

    data.forEach((row, index) => {
      const name = row.Name || row.name ? String(row.Name || row.name).trim() : '';
      const phone = row.Phone || row.phone ? String(row.Phone || row.phone).trim() : '';
      const method = row.Method || row.method ? String(row.Method || row.method).trim().toLowerCase() : 'whatsapp';
      const department = row.Department || row.department ? String(row.Department || row.department).trim() : 'Not selected';
      const subDepartment = row.SubDepartment || row.subDepartment ? String(row.SubDepartment || row.subDepartment).trim() : 'Not selected';
      const section = row.Section || row.section ? String(row.Section || row.section).trim() : 'Not selected';
      const designation = row.Designation || row.designation ? String(row.Designation || row.designation).trim() : 'Not selected';

      if (!name || !phone) {
        errors.push(`Row ${index + 2}: Missing name or phone`);
        return;
      }

      if (!phone.match(/^\+\d{10,15}$/)) {
        errors.push(`Row ${index + 2}: Invalid phone format (${phone})`);
        return;
      }

      if (!validMethods.includes(method)) {
        errors.push(`Row ${index + 2}: Invalid communication method (${method})`);
        return;
      }

      newCandidates.push({
        name,
        phone,
        method,
        status: 'Pending',
        response: null,
        interviewDate: null,
        interviewTime: null,
        department,
        subDepartment,
        section,
        designation,
        step: department === 'Not selected' ? 'welcome' : (subDepartment === 'Not selected' ? 'subDepartment' : (section === 'Not selected' && department !== 'E Department' ? 'section' : 'designation')),
        answers: [],
        correctAnswers: 0,
        currentQuestionIndex: 0
      });
    });

    if (newCandidates.length === 0) {
      console.error('[UPLOAD-CANDIDATES] No valid candidates found');
      return res.status(400).json({ error: 'No valid candidates found', details: errors });
    }

    if (candidates.length + newCandidates.length > 20) {
      console.error('[UPLOAD-CANDIDATES] Exceeds maximum candidate limit (20)');
      return res.status(400).json({ error: 'Adding candidates would exceed maximum limit of 20' });
    }

    candidates.push(...newCandidates);
    console.log('[UPLOAD-CANDIDATES] Added candidates:', JSON.stringify(newCandidates, null, 2));

    res.json({
      message: `Successfully added ${newCandidates.length} candidates`,
      candidates: newCandidates,
      errors: errors.length > 0 ? errors : undefined
    });
  } catch (error) {
    console.error('[UPLOAD-CANDIDATES] Error:', error);
    res.status(500).json({ error: 'Failed to process file' });
  }
});

// Handle manual multi-user submission
app.post('/send-manual', (req, res) => {
  const users = req.body;
  if (!users || !Array.isArray(users) || users.length === 0) {
    return res.status(400).json({ error: 'No users provided' });
  }

  const validMethods = ['voice', 'whatsapp'];
  const results = [];
  const errors = [];

  users.forEach((user, index) => {
    const name = user.name ? String(user.name).trim() : '';
    const phone = user.phone ? String(user.phone).trim() : '';
    const method = user.method ? String(user.method).trim().toLowerCase() : 'whatsapp';

    if (!name || !phone) {
      errors.push(`User ${index + 1}: Missing name or phone`);
      return;
    }

    if (!phone.match(/^\+\d{10,15}$/)) {
      errors.push(`User ${index + 1}: Invalid phone format (${phone})`);
      return;
    }

    if (!validMethods.includes(method)) {
      errors.push(`User ${index + 1}: Invalid communication method (${method})`);
      return;
    }

    const normalizedPhone = phone.startsWith('+') ? phone : `+${phone}`;
    candidates.push({
      name,
      phone: normalizedPhone,
      method,
      status: 'Pending',
      response: null,
      interviewDate: null,
      interviewTime: null,
      department: 'Not selected',
      subDepartment: 'Not selected',
      section: 'Not selected',
      designation: 'Not selected',
      step: 'welcome',
      answers: [],
      correctAnswers: 0,
      currentQuestionIndex: 0
    });

    results.push({
      name,
      phone: normalizedPhone,
      status: 'Added to candidates'
    });
  });

  if (candidates.length > 20) {
    candidates = candidates.slice(0, 20);
    errors.push('Candidate limit of 20 reached; excess candidates ignored');
  }

  console.log('[SEND-MANUAL] Added candidates:', JSON.stringify(results, null, 2));
  res.json({
    message: `Successfully added ${results.length} candidates`,
    candidates: results,
    errors: errors.length > 0 ? errors : undefined
  });
});

// Start interview endpoint
app.post('/start-interview', async (req, res) => {
  try {
    const pendingCandidates = candidates.filter(c => c.status === 'Pending');
    if (pendingCandidates.length === 0) {
      console.log('[START-INTERVIEW] No pending candidates');
      return res.json({ message: 'No pending candidates to process' });
    }

    for (const candidate of pendingCandidates) {
      const candidateIndex = candidates.findIndex(c => c.phone === candidate.phone);
      candidates[candidateIndex].status = 'In Progress';
      console.log(`[START-INTERVIEW] Processing ${candidate.method} for ${candidate.phone}`);

      try {
        if (candidate.method === 'voice') {
          await client.calls.create({
            url: `${NGROK_URL}/voice`,
            to: candidate.phone,
            from: twilioNumber,
            statusCallback: `${NGROK_URL}/call-status`,
            statusCallbackEvent: ['completed']
          });
          console.log(`[START-INTERVIEW] Voice call initiated to ${candidate.phone}`);
        } else {
          await client.messages.create({
            body: 'Welcome to the Jindal Saw Limited automated AI interview. Would you like to take an interview? 1 for Yes, 2 for No',
            from: whatsappNumber,
            to: `whatsapp:${candidate.phone}`
          });
          console.log(`[START-INTERVIEW] WhatsApp message sent to ${candidate.phone}`);
        }
      } catch (error) {
        console.error(`[START-INTERVIEW] Error processing ${candidate.method} for ${candidate.phone}:`, error);
        candidates[candidateIndex].status = 'Failed';
      }
    }

    res.json({ message: `Interview process initiated for ${pendingCandidates.length} candidates` });
  } catch (error) {
    console.error('[START-INTERVIEW] Error:', error);
    res.status(500).json({ error: 'Failed to start interviews' });
  }
});

// Voice response endpoint
app.post('/voice', (req, res) => {
  const VoiceResponse = twilio.twiml.VoiceResponse;
  const response = new VoiceResponse();
  console.log('[VOICE] Request body:', JSON.stringify(req.body, null, 2));

  try {
    const incomingPhone = req.body.To;
    const candidate = candidates.find(c => c.phone === incomingPhone);
    if (!candidate) {
      console.log(`[VOICE] No candidate found matching phone: ${incomingPhone}`);
      response.say('No active interview found. Goodbye.');
      res.type('text/xml');
      return res.send(response.toString());
    }

    candidate.status = 'In Progress';
    candidate.step = 'welcome';

    const gather = response.gather({
      numDigits: 1,
      action: '/welcome-response',
      method: 'POST',
      timeout: 5
    });
    gather.say('Welcome to the Jindal Saw Limited automated AI interview. Would you like to take an interview? Press 1 for Yes, 2 for No.');
  } catch (error) {
    console.error('[VOICE] Error:', error);
    response.say('An error occurred. Please try again later. Goodbye.');
  }

  res.type('text/xml');
  res.send(response.toString());
});

// Welcome response endpoint for voice
app.post('/welcome-response', (req, res) => {
  const VoiceResponse = twilio.twiml.VoiceResponse;
  const response = new VoiceResponse();
  console.log('[WELCOME-RESPONSE] Request body:', JSON.stringify(req.body, null, 2));

  try {
    const digit = req.body.Digits;
    const candidate = candidates.find(c => c.phone === req.body.From);
    if (!candidate || !digit || !['1', '2'].includes(digit)) {
      console.log('[WELCOME-RESPONSE] Invalid input or no candidate');
      response.say('Invalid input. Goodbye.');
      candidate && (candidate.status = 'Failed');
      res.type('text/xml');
      return res.send(response.toString());
    }

    if (digit === '1') {
      candidate.step = 'department';
      const gather = response.gather({
        numDigits: 1,
        action: '/select-department',
        method: 'POST',
        timeout: 5
      });
      gather.say(`Please select a department: ${departments.map((d, i) => `${i + 1}. ${d}`).join(', ')}.`);
    } else {
      candidate.status = 'Declined';
      candidate.step = 'done';
      response.say('Thank you for your response. Goodbye.');
    }
  } catch (error) {
    console.error('[WELCOME-RESPONSE] Error:', error);
    response.say('An error occurred. Please try again later. Goodbye.');
  }

  res.type('text/xml');
  res.send(response.toString());
});

// Department selection endpoint for voice
app.post('/select-department', (req, res) => {
  const VoiceResponse = twilio.twiml.VoiceResponse;
  const response = new VoiceResponse();
  console.log('[SELECT-DEPARTMENT] Request body:', JSON.stringify(req.body, null, 2));

  try {
    const digit = req.body.Digits;
    const candidate = candidates.find(c => c.phone === req.body.From);
    if (!candidate || !digit) {
      console.log('[SELECT-DEPARTMENT] Invalid input or no candidate');
      response.say('Invalid input. Goodbye.');
      candidate && (candidate.status = 'Failed');
      res.type('text/xml');
      return res.send(response.toString());
    }

    const deptIndex = parseInt(digit) - 1;
    if (deptIndex >= 0 && deptIndex < departments.length) {
      candidate.department = departments[deptIndex];
      candidate.step = 'subDepartment';
      console.log(`[SELECT-DEPARTMENT] Candidate ${candidate.phone} selected department: ${candidate.department}`);
      const gather = response.gather({
        numDigits: 1,
        action: '/select-subdepartment',
        method: 'POST',
        timeout: 5
      });
      gather.say(`You selected ${candidate.department}. Please select a sub-department: ${subDepartments[candidate.department].map((sd, i) => `${i + 1}. ${sd}`).join(', ')}.`);
    } else {
      const gather = response.gather({
        numDigits: 1,
        action: '/select-department',
        method: 'POST',
        timeout: 5
      });
      gather.say(`Please select a valid department: ${departments.map((d, i) => `${i + 1}. ${d}`).join(', ')}.`);
    }
  } catch (error) {
    console.error('[SELECT-DEPARTMENT] Error:', error);
    response.say('An error occurred. Please try again later. Goodbye.');
  }

  res.type('text/xml');
  res.send(response.toString());
});

// Sub-department selection endpoint for voice
app.post('/select-subdepartment', (req, res) => {
  const VoiceResponse = twilio.twiml.VoiceResponse;
  const response = new VoiceResponse();
  console.log('[SELECT-SUBDEPARTMENT] Request body:', JSON.stringify(req.body, null, 2));

  try {
    const digit = req.body.Digits;
    const candidate = candidates.find(c => c.phone === req.body.From);
    if (!candidate || !digit) {
      console.log('[SELECT-SUBDEPARTMENT] Invalid input or no candidate');
      response.say('Invalid input. Goodbye.');
      candidate && (candidate.status = 'Failed');
      res.type('text/xml');
      return res.send(response.toString());
    }

    const subDeptList = subDepartments[candidate.department];
    const subDeptIndex = parseInt(digit) - 1;
    if (subDeptIndex >= 0 && subDeptIndex < subDeptList.length) {
      candidate.subDepartment = subDeptList[subDeptIndex];
      candidate.step = candidate.department === 'E Department' ? 'designation' : 'section';
      console.log(`[SELECT-SUBDEPARTMENT] Candidate ${candidate.phone} selected sub-department: ${candidate.subDepartment}`);
      const gather = response.gather({
        numDigits: 1,
        action: candidate.department === 'E Department' ? '/select-designation' : '/select-section',
        method: 'POST',
        timeout: 5
      });
      const prompt = candidate.department === 'E Department'
        ? `You selected ${candidate.subDepartment}. Please select a designation: ${designations[candidate.department][candidate.subDepartment][candidate.subDepartment].map((d, i) => `${i + 1}. ${d}`).join(', ')}.`
        : `You selected ${candidate.subDepartment}. Please select a section: ${sections[candidate.department][candidate.subDepartment].map((s, i) => `${i + 1}. ${s}`).join(', ')}.`;
      gather.say(prompt);
    } else {
      const gather = response.gather({
        numDigits: '1',
        action: '/select-subdepartment',
        method: 'POST',
        timeout: 5
      });
      gather.say(`Please select a valid sub-department: ${subDeptList.map((sd, i) => `${i + 1}. ${sd}`).join(', ')}.`);
    }
  } catch (error) {
    console.error('[SELECT-SUBDEPARTMENT] Error:', error);
    response.say('An error occurred. Please try again later. Goodbye.');
  }

  res.type('text/xml');
  res.send(response.toString());
});

// Section selection endpoint for voice
app.post('/select-section', (req, res) => {
  const VoiceResponse = twilio.twiml.VoiceResponse;
  const response = new VoiceResponse();
  console.log('[SELECT-SECTION] Request body:', JSON.stringify(req.body, null, 2));

  try {
    const digit = req.body.Digits;
    const candidate = candidates.find(c => c.phone === req.body.From);
    if (!candidate || !digit) {
      console.log('[SELECT-SECTION] Invalid input or no candidate');
      response.say('Invalid input. Goodbye.');
      candidate && (candidate.status = 'Failed');
      res.type('text/xml');
      return res.send(response.toString());
    }

    const sectionList = sections[candidate.department][candidate.subDepartment];
    const sectionIndex = parseInt(digit) - 1;
    if (sectionIndex >= 0 && sectionIndex < sectionList.length) {
      candidate.section = sectionList[sectionIndex];
      candidate.step = 'designation';
      console.log(`[SELECT-SECTION] Candidate ${candidate.phone} selected section: ${candidate.section}`);
      const gather = response.gather({
        numDigits: 1,
        action: '/select-designation',
        method: 'POST',
        timeout: 5
      });
      gather.say(`You selected ${candidate.section}. Please select a designation: ${designations[candidate.department][candidate.subDepartment][candidate.section].map((d, i) => `${i + 1}. ${d}`).join(', ')}.`);
    } else {
      const gather = response.gather({
        numDigits: 1,
        action: '/select-section',
        method: 'POST',
        timeout: 5
      });
      gather.say(`Please select a valid section: ${sectionList.map((s, i) => `${i + 1}. ${s}`).join(', ')}.`);
    }
  } catch (error) {
    console.error('[SELECT-SECTION] Error:', error);
    response.say('An error occurred. Please try again later. Goodbye.');
  }

  res.type('text/xml');
  res.send(response.toString());
});

// Designation selection endpoint for voice
app.post('/select-designation', (req, res) => {
  const VoiceResponse = twilio.twiml.VoiceResponse;
  const response = new VoiceResponse();
  console.log('[SELECT-DESIGNATION] Request body:', JSON.stringify(req.body, null, 2));

  try {
    const digit = req.body.Digits;
    const candidate = candidates.find(c => c.phone === req.body.From);
    if (!candidate || !digit) {
      console.log('[SELECT-DESIGNATION] Invalid input or no candidate');
      response.say('Invalid input. Goodbye.');
      candidate && (candidate.status = 'Failed');
      res.type('text/xml');
      return res.send(response.toString());
    }

    const desigList = designations[candidate.department][candidate.subDepartment][candidate.department === 'E Department' ? candidate.subDepartment : candidate.section];
    const desigIndex = parseInt(digit) - 1;
    if (desigIndex >= 0 && desigIndex < desigList.length) {
      candidate.designation = desigList[desigIndex];
      candidate.step = 'schedule';
      console.log(`[SELECT-DESIGNATION] Candidate ${candidate.phone} selected designation: ${candidate.designation}`);
      const gather = response.gather({
        numDigits: 1,
        action: '/answer',
        method: 'POST',
        timeout: 5
      });
      gather.say(`You selected ${candidate.designation}. Would you like to schedule an interview? Press 1 for Yes, 2 for No.`);
    } else {
      const gather = response.gather({
        numDigits: 1,
        action: '/select-designation',
        method: 'POST',
        timeout: 5
      });
      gather.say(`Please select a valid designation: ${desigList.map((d, i) => `${i + 1}. ${d}`).join(', ')}.`);
    }
  } catch (error) {
    console.error('[SELECT-DESIGNATION] Error:', error);
    response.say('An error occurred. Please try again later. Goodbye.');
  }

  res.type('text/xml');
  res.send(response.toString());
});

// Voice answer endpoint
app.post('/answer', (req, res) => {
  const VoiceResponse = twilio.twiml.VoiceResponse;
  const response = new VoiceResponse();
  console.log('[ANSWER] Request body:', JSON.stringify(req.body, null, 2));

  try {
    const digit = req.body.Digits;
    const candidate = candidates.find(c => c.phone === req.body.From);
    if (!candidate || !digit || !['1', '2'].includes(digit)) {
      console.log('[ANSWER] Invalid input or no candidate');
      response.say('Invalid input. Goodbye.');
      candidate && (candidate.status = 'Failed');
      res.type('text/xml');
      return res.send(response.toString());
    }

    candidate.response = digit === '1' ? 'Yes' : 'No';
    console.log('[ANSWER] Candidate response:', candidate.response);

    if (candidate.response === 'Yes') {
      const gather = response.gather({
        numDigits: 1,
        action: '/select-date',
        method: 'POST',
        timeout: 5
      });
      gather.say('Please select your interview date. Press 1 for today, 2 for tomorrow, or 3 for the day after.');
    } else {
      response.say('Thank you for your response. Goodbye.');
      candidate.status = 'Declined';
    }
  } catch (error) {
    console.error('[ANSWER] Error:', error);
    response.say('An error occurred. Please try again later. Goodbye.');
  }

  res.type('text/xml');
  res.send(response.toString());
});

// Date selection endpoint for voice
app.post('/select-date', (req, res) => {
  const VoiceResponse = twilio.twiml.VoiceResponse;
  const response = new VoiceResponse();
  console.log('[SELECT-DATE] Request body:', JSON.stringify(req.body, null, 2));

  try {
    const digit = req.body.Digits;
    const candidate = candidates.find(c => c.phone === req.body.From);
    if (!candidate || !digit || !['1', '2', '3'].includes(digit)) {
      console.log('[SELECT-DATE] Invalid input or no candidate');
      response.say('Invalid input. Goodbye.');
      candidate && (candidate.status = 'Failed');
      res.type('text/xml');
      return res.send(response.toString());
    }

    const today = moment.tz('Asia/Kolkata');
    let selectedDate;
    if (digit === '1') {
      selectedDate = today;
    } else if (digit === '2') {
      selectedDate = today.clone().add(1, 'days');
    } else {
      selectedDate = today.clone().add(2, 'days');
    }
    candidate.interviewDate = selectedDate.format('YYYY-MM-DD');
    const availableSlots = getAvailableTimeSlots(candidate.interviewDate);
    if (availableSlots.length === 0) {
      response.say(`Sorry, no time slots are available for ${candidate.interviewDate}. Please select another date.`);
      const gather = response.gather({
        numDigits: 1,
        action: '/select-date',
        method: 'POST',
        timeout: 5
      });
      gather.say('Press 1 for today, 2 for tomorrow, or 3 for the day after.');
    } else {
      const gather = response.gather({
        numDigits: 1,
        action: '/select-time',
        method: 'POST',
        timeout: 5
      });
      gather.say(`Please select a time slot for ${candidate.interviewDate}. Available slots: ${availableSlots.map((slot, i) => `${i + 1}. ${slot}`).join(', ')}.`);
      candidate.step = 'time';
    }
  } catch (error) {
    console.error('[SELECT-DATE] Error:', error);
    response.say('An error occurred. Please try again later. Goodbye.');
  }

  res.type('text/xml');
  res.send(response.toString());
});

// Time selection endpoint for voice
app.post('/select-time', (req, res) => {
  const VoiceResponse = twilio.twiml.VoiceResponse;
  const response = new VoiceResponse();
  console.log('[SELECT-TIME] Request body:', JSON.stringify(req.body, null, 2));

  try {
    const digit = req.body.Digits;
    const candidate = candidates.find(c => c.phone === req.body.From);
    if (!candidate || !digit) {
      console.log('[SELECT-TIME] Invalid input or no candidate');
      response.say('Invalid input. Goodbye.');
      candidate && (candidate.status = 'Failed');
      res.type('text/xml');
      return res.send(response.toString());
    }

    if (candidate.department === 'Not selected' || candidate.subDepartment === 'Not selected' || (candidate.department !== 'E Department' && candidate.section === 'Not selected') || candidate.designation === 'Not selected') {
      console.error(`[SELECT-TIME] Invalid candidate data:`, JSON.stringify(candidate, null, 2));
      response.say('Invalid candidate data. Please start over. Goodbye.');
      candidate.status = 'Failed';
      res.type('text/xml');
      return res.send(response.toString());
    }

    const availableSlots = getAvailableTimeSlots(candidate.interviewDate);
    const timeIndex = parseInt(digit) - 1;
    if (timeIndex >= 0 && timeIndex < availableSlots.length) {
      const selectedSlot = availableSlots[timeIndex];
      if (!bookedTimeSlots[candidate.interviewDate]?.includes(selectedSlot)) {
        candidate.interviewTime = selectedSlot;
        if (!bookedTimeSlots[candidate.interviewDate]) {
          bookedTimeSlots[candidate.interviewDate] = [];
        }
        bookedTimeSlots[candidate.interviewDate].push(candidate.interviewTime);
        const nextDesignation = getNextDesignation(candidate.department, candidate.subDepartment, candidate.section || candidate.subDepartment, candidate.designation);
        candidate.status = `Scheduled for ${nextDesignation} (current: ${candidate.designation})`;
        console.log(`[SELECT-TIME] Candidate ${candidate.phone} selected time slot: ${candidate.interviewTime}`);
        scheduleInterview(candidate);
        response.say(`Your interview is scheduled on ${candidate.interviewDate} at ${candidate.interviewTime}. Thank you!`);
        client.messages.create({
          body: `Your interview is scheduled on ${candidate.interviewDate} at ${candidate.interviewTime}.`,
          from: whatsappNumber,
          to: `whatsapp:${candidate.phone}`
        });
      } else {
        const gather = response.gather({
          numDigits: 1,
          action: '/select-time',
          method: 'POST',
          timeout: 5
        });
        gather.say(`Sorry, that time slot is taken. Please select another: ${availableSlots.map((slot, i) => `${i + 1}. ${slot}`).join(', ')}.`);
      }
    } else {
      const gather = response.gather({
        numDigits: 1,
        action: '/select-time',
        method: 'POST',
        timeout: 5
      });
      gather.say(`Please select a valid time slot: ${availableSlots.map((slot, i) => `${i + 1}. ${slot}`).join(', ')}.`);
    }
  } catch (error) {
    console.error('[SELECT-TIME] Error:', error);
    response.say('An error occurred. Please try again later. Goodbye.');
  }

  res.type('text/xml');
  res.send(response.toString());
});

// Voice question endpoint
app.post('/voice-question', (req, res) => {
  const VoiceResponse = twilio.twiml.VoiceResponse;
  const response = new VoiceResponse();
  console.log('[VOICE-QUESTION] Request body:', JSON.stringify(req.body, null, 2));

  try {
    const candidate = candidates.find(c => c.phone === req.body.From);
    if (!candidate) {
      console.log('[VOICE-QUESTION] No candidate in interview');
      response.say('No active interview found. Goodbye.');
      res.type('text/xml');
      return res.send(response.toString());
    }

    const questions = questionBank[candidate.department][candidate.subDepartment][candidate.section || candidate.subDepartment][candidate.designation] || [];
    const questionIndex = candidate.currentQuestionIndex || 0;

    if (questionIndex >= questions.length) {
      candidate.status = 'Interview Completed';
      response.say('Thank you for completing the interview. Goodbye.');
      res.type('text/xml');
      return res.send(response.toString());
    }

    const gather = response.gather({
      numDigits: 1,
      action: '/voice-answer',
      method: 'POST',
      timeout: 5
    });
    gather.say(questions[questionIndex].question);
  } catch (error) {
    console.error('[VOICE-QUESTION] Error:', error);
    response.say('An error occurred. Please try again later. Goodbye.');
  }

  res.type('text/xml');
  res.send(response.toString());
});

// Voice answer endpoint
app.post('/voice-answer', (req, res) => {
  const VoiceResponse = twilio.twiml.VoiceResponse;
  const response = new VoiceResponse();
  console.log('[VOICE-ANSWER] Request body:', JSON.stringify(req.body, null, 2));

  try {
    const digit = req.body.Digits;
    const candidate = candidates.find(c => c.phone === req.body.From);
    if (!candidate || !digit) {
      console.log('[VOICE-ANSWER] Invalid input or no candidate');
      response.say('Invalid input. Goodbye.');
      candidate && (candidate.status = 'Failed');
      res.type('text/xml');
      return res.send(response.toString());
    }

    const questions = questionBank[candidate.department][candidate.subDepartment][candidate.section || candidate.subDepartment][candidate.designation] || [];
    const questionIndex = candidate.currentQuestionIndex || 0;
    candidate.answers.push(digit);
    if (digit === questions[questionIndex].correctAnswer) {
      candidate.correctAnswers = (candidate.correctAnswers || 0) + 1;
    }

    candidate.currentQuestionIndex = questionIndex + 1;
    if (candidate.currentQuestionIndex >= questions.length) {
      candidate.status = 'Interview Completed';
      response.say('Thank you for completing the interview. Goodbye.');
    } else {
      const gather = response.gather({
        numDigits: 1,
        action: '/voice-question',
        method: 'POST',
        timeout: 5
      });
      gather.say(questions[candidate.currentQuestionIndex].question);
    }
  } catch (error) {
    console.error('[VOICE-ANSWER] Error:', error);
    response.say('An error occurred. Please try again later. Goodbye.');
  }

  res.type('text/xml');
  res.send(response.toString());
});

// Call status callback
app.post('/call-status', (req, res) => {
  console.log('[CALL-STATUS] Request body:', JSON.stringify(req.body, null, 2));
  try {
    const incomingPhone = req.body.To || req.body.From;
    const candidate = candidates.find(c => c.phone === incomingPhone);
    if (candidate) {
      const candidateIndex = candidates.findIndex(c => c.phone === candidate.phone);
      candidates[candidateIndex].status = req.body.CallStatus === 'completed' ? 'Interview Completed' : 'Failed';
      console.log(`[CALL-STATUS] Updated candidate ${candidate.phone} to ${candidates[candidateIndex].status}`);
    }
    res.sendStatus(200);
  } catch (error) {
    console.error('[CALL-STATUS] Error:', error);
    res.sendStatus(500);
  }
});

// WhatsApp webhook
app.post('/webhook', async (req, res) => {
  const from = req.body.From;
  const message = req.body.Body.trim().toLowerCase();
  console.log(`[WEBHOOK] Received from ${from}: ${message}`);

  let candidate = candidates.find(c => c.phone === from.replace('whatsapp:', '') && c.method === 'whatsapp');
  if (!candidate) {
    candidate = {
      name: 'Unknown',
      phone: from.replace('whatsapp:', ''),
      method: 'whatsapp',
      status: 'Started via WhatsApp',
      response: null,
      interviewDate: null,
      interviewTime: null,
      department: 'Not selected',
      subDepartment: 'Not selected',
      section: 'Not selected',
      designation: 'Not selected',
      step: 'welcome',
      answers: [],
      correctAnswers: 0,
      currentQuestionIndex: 0
    };
    candidates.push(candidate);
  }

  if (['Scheduled', 'Declined', 'Failed'].includes(candidate.status) && candidate.step === 'done') {
    console.log(`[WEBHOOK] Resuming candidate ${from} from ${candidate.status}`);
    candidate.status = 'In Progress';
    if (candidate.interviewDate) {
      candidate.step = 'time';
    } else if (candidate.designation !== 'Not selected') {
      candidate.step = 'schedule';
    } else if (candidate.section !== 'Not selected') {
      candidate.step = 'designation';
    } else if (candidate.subDepartment !== 'Not selected') {
      candidate.step = candidate.department === 'E Department' ? 'designation' : 'section';
    } else if (candidate.department !== 'Not selected') {
      candidate.step = 'subDepartment';
    } else if (candidate.name !== 'Unknown') {
      candidate.step = 'department';
    } else {
      candidate.step = 'welcome';
    }
  }

  let response = '';

  switch (candidate.step) {
    case 'welcome':
      if (message === '1' || message === 'yes') {
        candidate.status = 'In Progress';
        response = 'Please provide your name.';
        candidate.step = 'name';
      } else if (message === '2' || message === 'no') {
        candidate.status = 'Declined';
        response = 'Thank you for your response. Goodbye.';
        candidate.step = 'done';
      } else {
        response = 'Welcome to the Jindal Saw Limited automated AI interview. Would you like to take an interview? 1 for Yes, 2 for No';
      }
      break;
    case 'name':
      candidate.name = message.charAt(0).toUpperCase() + message.slice(1);
      candidate.status = 'Name provided';
      response = `Hello ${candidate.name}, please select a department:\n${departments.map((d, i) => `${i + 1}. ${d}`).join('\n')}`;
      candidate.step = 'department';
      break;
    case 'department':
      const deptIndex = parseInt(message) - 1;
      if (deptIndex >= 0 && deptIndex < departments.length) {
        candidate.department = departments[deptIndex];
        candidate.status = 'Department selected';
        response = `You selected ${candidate.department}.\nPlease select a sub-department:\n${subDepartments[candidate.department].map((sd, i) => `${i + 1}. ${sd}`).join('\n')}`;
        candidate.step = 'subDepartment';
      } else {
        response = `Please select a valid department:\n${departments.map((d, i) => `${i + 1}. ${d}`).join('\n')}`;
      }
      break;
    case 'subDepartment':
      const subDeptList = subDepartments[candidate.department];
      const subDeptIndex = parseInt(message) - 1;
      if (subDeptIndex >= 0 && subDeptIndex < subDeptList.length) {
        candidate.subDepartment = subDeptList[subDeptIndex];
        candidate.status = 'Sub-department selected';
        candidate.step = candidate.department === 'E Department' ? 'designation' : 'section';
        response = candidate.department === 'E Department'
          ? `You selected ${candidate.subDepartment}.\nPlease select a designation:\n${designations[candidate.department][candidate.subDepartment][candidate.subDepartment].map((d, i) => `${i + 1}. ${d}`).join('\n')}`
          : `You selected ${candidate.subDepartment}.\nPlease select a section:\n${sections[candidate.department][candidate.subDepartment].map((s, i) => `${i + 1}. ${s}`).join('\n')}`;
      } else {
        response = `Please select a valid sub-department:\n${subDeptList.map((sd, i) => `${i + 1}. ${sd}`).join('\n')}`;
      }
      break;
    case 'section':
      const sectionList = sections[candidate.department][candidate.subDepartment];
      const sectionIndex = parseInt(message) - 1;
      if (sectionIndex >= 0 && sectionIndex < sectionList.length) {
        candidate.section = sectionList[sectionIndex];
        candidate.status = 'Section selected';
        response = `You selected ${candidate.section}.\nPlease select a designation:\n${designations[candidate.department][candidate.subDepartment][candidate.section].map((d, i) => `${i + 1}. ${d}`).join('\n')}`;
        candidate.step = 'designation';
      } else {
        response = `Please select a valid section:\n${sectionList.map((s, i) => `${i + 1}. ${s}`).join('\n')}`;
      }
      break;
    case 'designation':
      const desigList = designations[candidate.department][candidate.subDepartment][candidate.department === 'E Department' ? candidate.subDepartment : candidate.section];
      const desigIndex = parseInt(message) - 1;
      if (desigIndex >= 0 && desigIndex < desigList.length) {
        candidate.designation = desigList[desigIndex]; 
        candidate.status = 'Designation selected';
        response = `You selected ${candidate.designation}.\nWould you like to schedule an interview?\n1. Yes\n2. No`;
        candidate.step = 'schedule';
      } else {
        response = `Please select a valid designation:\n${desigList.map((d, i) => `${i + 1}. ${d}`).join('\n')}`;
      }
      break;
    case 'schedule':
      if (message === '1' || message === 'yes') {
        candidate.status = 'Scheduling interview';
        response = 'Please select your interview date:\n1. Today\n2. Tomorrow\n3. Day after';
        candidate.step = 'date';
      } else if (message === '2' || message === 'no') {
        candidate.status = 'Declined';
        response = 'Thank you! If you wish to schedule an interview later, please contact again.';
        candidate.step = 'done';
      } else {
        response = 'Please select 1 (Yes) or 2 (No).';
      }
      break;
    case 'date':
      if (['1', '2', '3'].includes(message)) {
        const today = moment.tz('Asia/Kolkata');
        let selectedDate;
        if (message === '1') {
          selectedDate = today;
        } else if (message === '2') {
          selectedDate = today.clone().add(1, 'days');
        } else {
          selectedDate = today.clone().add(2, 'days');
        }
        candidate.interviewDate = selectedDate.format('YYYY-MM-DD');
        candidate.status = 'Date selected';
        const availableSlots = getAvailableTimeSlots(candidate.interviewDate);
        if (availableSlots.length === 0) {
          response = `Sorry, no time slots are available for ${candidate.interviewDate}. Please select another date:\n1. Today\n2. Tomorrow\n3. Day after`;
          candidate.step = 'date';
        } else {
          response = `Please select a time slot for ${candidate.interviewDate}:\n${availableSlots.map((slot, i) => `${i + 1}. ${slot}`).join('\n')}`;
          candidate.step = 'time';
        }
      } else {
        response = 'Please select a valid date: 1 for today, 2 for tomorrow, or 3 for the day after.';
      }
      break;
    case 'time':
      const availableSlots = getAvailableTimeSlots(candidate.interviewDate);
      const timeIndex = parseInt(message) - 1;
      if (timeIndex >= 0 && timeIndex < availableSlots.length) {
        const selectedSlot = availableSlots[timeIndex];
        if (!bookedTimeSlots[candidate.interviewDate]?.includes(selectedSlot)) {
          candidate.interviewTime = selectedSlot;
          if (!bookedTimeSlots[candidate.interviewDate]) {
            bookedTimeSlots[candidate.interviewDate] = [];
          }
          bookedTimeSlots[candidate.interviewDate].push(candidate.interviewTime);
          const nextDesignation = getNextDesignation(candidate.department, candidate.subDepartment, candidate.section || candidate.subDepartment, candidate.designation);
          candidate.status = `Scheduled for ${nextDesignation} (current: ${candidate.designation})`;
          response = `Your interview is scheduled on ${candidate.interviewDate} at ${candidate.interviewTime}. Your current designation is ${candidate.designation} and you are scheduled for ${nextDesignation}. Please be on time. Thank you!`;
          scheduleInterview(candidate);
          candidate.step = 'question';
        } else {
          response = `Sorry, this time slot is no longer available. Please select another slot:\n${getAvailableTimeSlots(candidate.interviewDate).map((slot, i) => `${i + 1}. ${slot}`).join('\n')}`;
        }
      } else {
        response = `Please select a valid time slot:\n${availableSlots.map((slot, i) => `${i + 1}. ${slot}`).join('\n')}`;
      }
      break;
    case 'question':
      response = 'Your interview questions will be asked via a voice call at the scheduled time. Please be prepared to receive the call.';
      candidate.step = 'done';
      break;
    default:
      response = 'Invalid input. Please start over.';
      candidate.status = 'Failed';
      candidate.step = 'done';
  }

  try {
    if (response) {
      await client.messages.create({
        body: response,
        from: whatsappNumber,
        to: from
      });
      console.log(`[WEBHOOK] Response sent to ${from}: ${response}`);
    }
    res.status(200).send('OK');
  } catch (error) {
    console.error(`[WEBHOOK] Error sending response to ${from}:`, error);
    res.status(500).send('Error processing webhook');
  }
});

// Results endpoint
app.get('/results', (req, res) => {
  try {
    const results = candidates.map(candidate => ({
      name: candidate.name,
      phone: candidate.phone,
      method: candidate.method,
      status: candidate.status,
      interviewDate: candidate.interviewDate,
      interviewTime: candidate.interviewTime,
      department: candidate.department,
      subDepartment: candidate.subDepartment,
      section: candidate.section,
      designation: candidate.designation,
      score: candidate.correctAnswers || 0,
      totalQuestions: questionBank[candidate.department]?.[candidate.subDepartment]?.[candidate.section || candidate.subDepartment]?.[candidate.designation]?.length || 0
    }));
    console.log('[RESULTS] Returning results:', JSON.stringify(results, null, 2));
    res.json(results);
  } catch (error) {
    console.error('[RESULTS] Error:', error);
    res.status(500).json({ error: 'Failed to fetch results' });
  }
});

// Periodic result fetching every 30 seconds
function fetchResultsPeriodically() {
  try {
    const results = candidates.map(candidate => ({
      name: candidate.name,
      phone: candidate.phone,
      method: candidate.method,
      status: candidate.status,
      interviewDate: candidate.interviewDate,
      interviewTime: candidate.interviewTime,
      department: candidate.department,
      subDepartment: candidate.subDepartment,
      section: candidate.section,
      designation: candidate.designation,
      score: candidate.correctAnswers || 0,
      totalQuestions: questionBank[candidate.department]?.[candidate.subDepartment]?.[candidate.section || candidate.subDepartment]?.[candidate.designation]?.length || 0
    }));
    console.log('[PERIODIC-RESULTS] Fetched results at', new Date().toISOString(), ':', JSON.stringify(results, null, 2));
  } catch (error) {
    console.error('[PERIODIC-RESULTS] Error fetching results:', error);
  }
}

// Start the server and periodic result fetching
const PORT = process.env.PORT || 3000;
app.listen(PORT, '0.0.0.0', () => {
  console.log(`Server is running on port ${PORT}`);
  setInterval(fetchResultsPeriodically, 30 * 1000);
  fetchResultsPeriodically();
});