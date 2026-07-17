require('dotenv').config();
const path = require('path');
const express = require('express');
const { GoogleSpreadsheet } = require('google-spreadsheet');
const { JWT } = require('google-auth-library');
const creds = require(process.env.GOOGLE_PRIVATE_KEY_PATH);

const app = express();
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static('public'));

app.use((req, res, next) => {
    console.log(`[REQUEST] ${req.method} ${req.url}`);
    next();
});

app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));

const auth = new JWT({
    email: creds.client_email,
    key: creds.private_key,
    scopes: [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ],
});

const MASTER_TAB = 'Clients Tracker';

const ADMIN_ACCOUNTS = {
    'maittomgr@gmail.com': { name: 'Manigandaraja' },
    'mythkuki11@gmail.com': { name: 'Admin' }
};

const fs = require('fs');

const ROSTER_FILE = path.join(__dirname, 'roster.json');
const TEAM_ACCOUNTS_FILE = path.join(__dirname, 'team_accounts.json');

let TEAM_ACCOUNTS = {
    'guhanredaye@gmail.com': { name: 'Mani' },
    'ccmmonikam@gmail.com': { name: 'Monika' },
    'ccmkavitham@gmail.com': { name: 'Kavitha' },
    'ccmabinayasri@gmail.com': { name: 'Abinayasri' },
    'dharunika@example.com': { name: 'Dharunika' },
    'dharani@example.com': { name: 'Dharani' },
    'ccmsowndark@gmail.com': { name: 'Sowndar' },
    'manoranjan@example.com': { name: 'Manoranjan' }
};

let ROSTER = [
    { name: 'Manigandaraja', team: 'Creative' },
    { name: 'Monika', team: 'Creative' },
    { name: 'Kavitha', team: 'Production' },
    { name: 'Abinayasri', team: 'Production' },
    { name: 'Dharunika', team: 'Production' },
    { name: 'Dharani', team: 'Production' },
    { name: 'Sowndar', team: 'Editor' },
    { name: 'Manoranjan', team: 'Editor' }
];

try {
    if (fs.existsSync(TEAM_ACCOUNTS_FILE)) {
        TEAM_ACCOUNTS = JSON.parse(fs.readFileSync(TEAM_ACCOUNTS_FILE, 'utf8'));
    } else {
        fs.writeFileSync(TEAM_ACCOUNTS_FILE, JSON.stringify(TEAM_ACCOUNTS, null, 2), 'utf8');
    }

    if (fs.existsSync(ROSTER_FILE)) {
        ROSTER = JSON.parse(fs.readFileSync(ROSTER_FILE, 'utf8'));
    } else {
        fs.writeFileSync(ROSTER_FILE, JSON.stringify(ROSTER, null, 2), 'utf8');
    }
} catch (e) {
    console.error('Error loading roster/accounts configuration:', e);
}

const ADMIN_EMAILS = Object.keys(ADMIN_ACCOUNTS);
let TEAM_EMAILS = {};
function refreshTeamEmails() {
    TEAM_EMAILS = {};
    for (const email in TEAM_ACCOUNTS) {
        TEAM_EMAILS[email] = TEAM_ACCOUNTS[email].name;
    }
}
refreshTeamEmails();

const PASSWORDS_FILE = path.join(__dirname, 'passwords.json');
let userPasswords = {};
try {
    if (fs.existsSync(PASSWORDS_FILE)) {
        userPasswords = JSON.parse(fs.readFileSync(PASSWORDS_FILE, 'utf8'));
    }
} catch (e) {
    console.error('Error reading passwords.json:', e);
}

function saveUserPassword(email, password) {
    userPasswords[email.toLowerCase().trim()] = password;
    try {
        fs.writeFileSync(PASSWORDS_FILE, JSON.stringify(userPasswords, null, 2), 'utf8');
    } catch (e) {
        console.error('Error writing to passwords.json:', e);
    }
}

function getEmailFromCookie(req) {
    const cookieHeader = req.headers.cookie || '';
    const match = cookieHeader.match(/user_email=([^;]+)/);
    return match ? decodeURIComponent(match[1]) : null;
}

// Authentication Middleware
function requireAuth(req, res, next) {
    const email = getEmailFromCookie(req);
    if (!email) {
        if (req.xhr || req.headers.accept?.includes('json')) {
            return res.status(401).json({ error: 'Unauthorized. Please log in.' });
        }
        return res.redirect('/login');
    }

    const emailLc = email.toLowerCase().trim();
    const isAdmin = ADMIN_EMAILS.map(e => e.toLowerCase()).includes(emailLc);
    const matchedKey = Object.keys(TEAM_EMAILS).find(k => k.toLowerCase() === emailLc);
    const isMember = !!matchedKey;

    if (!isAdmin && !isMember) {
        res.setHeader('Set-Cookie', 'user_email=; Path=/; HttpOnly; Max-Age=0');
        if (req.xhr || req.headers.accept?.includes('json')) {
            return res.status(403).json({ error: 'Forbidden. Access restricted.' });
        }
        return res.redirect('/login');
    }

    req.userEmail = email;
    req.isAdmin = isAdmin;
    req.userName = isAdmin ? 'Admin' : TEAM_EMAILS[matchedKey];
    next();
}

async function getSheet(tabName) {
    const doc = new GoogleSpreadsheet(process.env.SHEET_ID, auth);
    await doc.loadInfo();
    const sheet = doc.sheetsByTitle[tabName];
    if (tabName === MASTER_TAB) {
        await sheet.loadHeaderRow();
        const required = [
            'Script & Storyboard Status - Finished Date',
            'Image & Video Generation Status - Finished Date',
            'Editing Status - Finished Date'
        ];
        let changed = false;
        const newHeaders = [...sheet.headerValues];
        required.forEach(h => {
            if (!newHeaders.includes(h)) {
                newHeaders.push(h);
                changed = true;
            }
        });
        if (changed) {
            await sheet.setHeaderRow(newHeaders);
        }
    }
    return sheet;
}

function computeClients(tasks) {
    const counts = {};
    tasks.forEach(t => {
        const name = t['Client Name'];
        if (!name) return;
        counts[name] = (counts[name] || 0) + 1;
    });
    return Object.keys(counts).map(name => ({ name, count: counts[name] }));
}

async function findRowByVideoTitle(sheet, videoTitle) {
    const rows = await sheet.getRows();
    return rows.find(r => {
        const obj = r.toObject();
        return obj['Video Title'] === videoTitle;
    });
}

// ---------- Authentication Routes ----------
app.get('/login', (req, res) => {
    res.render('login', { error: null });
});

app.post('/login', (req, res) => {
    const email = (req.body.email || '').trim().toLowerCase();
    const password = (req.body.password || '').trim();

    if (!email || !password) {
        return res.render('login', { error: 'Email and Password are required.' });
    }

    const adminKey = Object.keys(ADMIN_ACCOUNTS).find(k => k.toLowerCase() === email);
    const memberKey = Object.keys(TEAM_ACCOUNTS).find(k => k.toLowerCase() === email);

    let accountName = null;
    if (adminKey) {
        accountName = ADMIN_ACCOUNTS[adminKey].name;
    } else if (memberKey) {
        accountName = TEAM_ACCOUNTS[memberKey].name;
    }

    if (!accountName) {
        return res.render('login', { error: `Access restricted. ${email} is not authorized.` });
    }

    const savedPassword = userPasswords[email];
    if (savedPassword) {
        if (savedPassword !== password) {
            return res.render('login', { error: 'Invalid password.' });
        }
    } else {
        // First time login! Set password
        saveUserPassword(email, password);
    }

    res.setHeader('Set-Cookie', `user_email=${encodeURIComponent(email)}; Path=/; HttpOnly; Max-Age=${30 * 24 * 60 * 60}`);
    res.redirect('/');
});

app.get('/logout', (req, res) => {
    res.setHeader('Set-Cookie', 'user_email=; Path=/; HttpOnly; Max-Age=0');
    res.redirect('/login');
});

// ---------- Dashboard page (server-rendered) ----------
app.get('/', requireAuth, async (req, res) => {
    try {
        let currentEmail = req.userEmail;
        let currentName = req.userName;
        let role = req.isAdmin ? 'admin' : 'member';

        if (req.isAdmin) {
            // Admins can override view using ?user=Name
            const requestedUser = (req.query.user || '').trim();
            if (requestedUser && requestedUser.toLowerCase() !== 'admin') {
                const matchedPerson = ROSTER.find(p => p.name.toLowerCase() === requestedUser.toLowerCase());
                if (matchedPerson) {
                    role = 'member';
                    currentName = matchedPerson.name;
                    for (const key in TEAM_EMAILS) {
                        if (TEAM_EMAILS[key].toLowerCase() === currentName.toLowerCase()) {
                            currentEmail = key;
                            break;
                        }
                    }
                }
            }
        }

        const sheet = await getSheet(MASTER_TAB);
        const rows = await sheet.getRows();
        const tasks = rows.map(r => r.toObject());

        res.render('index', {
            dataJson: JSON.stringify(tasks),
            rosterJson: JSON.stringify(ROSTER),
            clientsJson: JSON.stringify(computeClients(tasks)),
            roster: ROSTER,
            role,
            personName: currentName,
            userEmail: currentEmail,
            generatedAt: new Date().toLocaleString(),
            teamAccountsJson: JSON.stringify(TEAM_ACCOUNTS)
        });
    } catch (err) {
        console.error(err);
        res.status(500).send('Error loading dashboard: ' + err.message);
    }
});

// ---------- Read-only API (kept for the earlier MVP table / external use) ----------
app.get('/api/tasks', requireAuth, async (req, res) => {
    try {
        const sheet = await getSheet(MASTER_TAB);
        const rows = await sheet.getRows();
        res.json(rows.map(r => r.toObject()));
    } catch (err) {
        console.error(err);
        res.status(500).json({ error: err.message });
    }
});

app.patch('/api/tasks/:rowIndex', requireAuth, async (req, res) => {
    try {
        const sheet = await getSheet(MASTER_TAB);
        const rows = await sheet.getRows();
        const row = rows[req.params.rowIndex];
        if (!row) return res.status(404).json({ error: 'Row not found' });
        Object.assign(row, req.body);
        await row.save();
        res.json({ success: true });
    } catch (err) {
        console.error(err);
        res.status(500).json({ error: err.message });
    }
});

// ---------- Write-back actions used by the dashboard's My Tasks / Approvals / Assign tabs ----------
app.post('/api/startProgress', requireAuth, async (req, res) => {
    try {
        const { videoTitle, stageKey } = req.body;
        const statusFieldByStage = {
            script: 'Script & Storyboard Status',
            prod: 'Image & Video Generation Status',
            edit: 'Editing Status',
        };
        const roleFieldByStage = {
            script: 'Creative Team',
            prod: 'Production Team',
            edit: 'Editor',
        };
        const field = statusFieldByStage[stageKey];
        if (!field) return res.status(400).json({ error: 'Unknown stage: ' + stageKey });

        const sheet = await getSheet(MASTER_TAB);
        const row = await findRowByVideoTitle(sheet, videoTitle);
        if (!row) return res.status(404).json({ error: 'Video not found: ' + videoTitle });

        const assignedPerson = row.toObject()[roleFieldByStage[stageKey]];
        if (!req.isAdmin && assignedPerson !== req.userName) {
            return res.status(403).json({ error: `This task is assigned to ${assignedPerson || 'nobody'}, not you.` });
        }

        row.set(field, 'In progress');
        await row.save();
        res.json({ ok: true });
    } catch (err) {
        console.error(err);
        res.status(500).json({ error: err.message });
    }
});

app.post('/api/markComplete', requireAuth, async (req, res) => {
    try {
        const { videoTitle, stageKey } = req.body;
        const statusFieldByStage = {
            script: 'Script & Storyboard Status',
            prod: 'Image & Video Generation Status',
            edit: 'Editing Status',
        };
        const roleFieldByStage = {
            script: 'Creative Team',
            prod: 'Production Team',
            edit: 'Editor',
        };
        const field = statusFieldByStage[stageKey];
        if (!field) return res.status(400).json({ error: 'Unknown stage: ' + stageKey });

        const sheet = await getSheet(MASTER_TAB);
        const row = await findRowByVideoTitle(sheet, videoTitle);
        if (!row) return res.status(404).json({ error: 'Video not found: ' + videoTitle });

        const assignedPerson = row.toObject()[roleFieldByStage[stageKey]];
        if (!req.isAdmin && assignedPerson !== req.userName) {
            return res.status(403).json({ error: `This task is assigned to ${assignedPerson || 'nobody'}, not you.` });
        }

        const pad = (n) => String(n).padStart(2, '0');
        const d = new Date();
        const dateStr = pad(d.getDate()) + '/' + pad(d.getMonth() + 1) + '/' + d.getFullYear();
        const timeStr = pad(d.getHours()) + ':' + pad(d.getMinutes());
        const timestamp = `${dateStr} ${timeStr}`;

        const finishedDateFieldByStage = {
            script: 'Script & Storyboard Status - Finished Date',
            prod: 'Image & Video Generation Status - Finished Date',
            edit: 'Editing Status - Finished Date',
        };
        const finishedField = finishedDateFieldByStage[stageKey];

        row.set(field, 'Completed');
        if (finishedField) {
            row.set(finishedField, timestamp);
        }
        await row.save();
        res.json({ ok: true });
    } catch (err) {
        console.error(err);
        res.status(500).json({ error: err.message });
    }
});

app.post('/api/approveStage', requireAuth, async (req, res) => {
    try {
        if (!req.isAdmin) {
            return res.status(403).json({ error: 'Only the admin can approve tasks.' });
        }
        const { videoTitle, stageKey, decision, comment } = req.body;
        const fieldsByStage = {
            script: { status: 'Script & Storyboard Status', approval: 'Script & Storyboard Approval', message: 'Script & Storyboard Message' },
            prod: { status: 'Image & Video Generation Status', approval: 'Image & Video Generation Approval', message: 'Image & Video Generation Message' },
            edit: { status: 'Editing Status', approval: 'Editing Approval', message: 'Editing Message' },
        };
        const fields = fieldsByStage[stageKey];
        if (!fields) return res.status(400).json({ error: 'Unknown stage: ' + stageKey });

        const sheet = await getSheet(MASTER_TAB);
        const row = await findRowByVideoTitle(sheet, videoTitle);
        if (!row) return res.status(404).json({ error: 'Video not found: ' + videoTitle });

        if (decision === 'approve') {
            row.set(fields.approval, 'Approved');
            row.set(fields.status, 'Completed');
            row.set(fields.message, '');
        } else {
            row.set(fields.approval, 'Changes requested');
            row.set(fields.status, 'Changes requested');
            row.set(fields.message, comment || '');
        }
        await row.save();
        res.json({ ok: true });
    } catch (err) {
        console.error(err);
        res.status(500).json({ error: err.message });
    }
});

app.post('/api/setGateStatus', requireAuth, async (req, res) => {
    try {
        if (!req.isAdmin) {
            return res.status(403).json({ error: 'Only the admin can update this.' });
        }
        const { videoTitle, columnName, value } = req.body;
        const allowedColumns = ['Final Internal Approval', 'Client Approval', 'Broadcast Status'];
        if (!allowedColumns.includes(columnName)) return res.status(400).json({ error: 'Unknown column: ' + columnName });

        const sheet = await getSheet(MASTER_TAB);
        const row = await findRowByVideoTitle(sheet, videoTitle);
        if (!row) return res.status(404).json({ error: 'Video not found: ' + videoTitle });

        row.set(columnName, value);
        await row.save();
        res.json({ ok: true });
    } catch (err) {
        console.error(err);
        res.status(500).json({ error: err.message });
    }
});

app.post('/api/assignPerson', requireAuth, async (req, res) => {
    try {
        if (!req.isAdmin) {
            return res.status(403).json({ error: 'Only the admin can assign tasks.' });
        }
        const { videoTitle, roleCol, personName } = req.body;
        const allowedColumns = [
            'Creative Team', 'Production Team', 'Editor', 'Priority', 'Deadline',
            'Creative Link', 'Production Link', 'Editor Link'
        ];
        if (!allowedColumns.includes(roleCol)) return res.status(400).json({ error: 'Unknown role column: ' + roleCol });

        const sheet = await getSheet(MASTER_TAB);
        const row = await findRowByVideoTitle(sheet, videoTitle);
        if (!row) return res.status(404).json({ error: 'Video not found: ' + videoTitle });

        row.set(roleCol, personName || '');
        await row.save();
        res.json({ ok: true });
    } catch (err) {
        console.error(err);
        res.status(500).json({ error: err.message });
    }
});

app.post('/api/createTask', requireAuth, async (req, res) => {
    try {
        if (!req.isAdmin) {
            return res.status(403).json({ error: 'Only the admin can add new tasks.' });
        }
        const { clientName, videoTitle, priority, deadline, creative, production, editor } = req.body;
        if (!clientName || !videoTitle) {
            return res.status(400).json({ error: 'Client Name and Video Title are required.' });
        }

        const sheet = await getSheet(MASTER_TAB);
        const rows = await sheet.getRows();
        const exists = rows.some(r => r.get('Video Title')?.toLowerCase() === videoTitle.toLowerCase());
        if (exists) {
            return res.status(400).json({ error: `A video with title "${videoTitle}" already exists.` });
        }

        const newRowData = {
            'Client Name': clientName,
            'Video Title': videoTitle,
            'Priority': priority || 'Normal',
            'Deadline': deadline || '',
            'Creative Team': creative || '',
            'Script & Storyboard Status': 'Not started',
            'Script & Storyboard Approval': 'Not yet approved',
            'Script & Storyboard Message': '',
            'Script & Storyboard Status - Finished Date': '',
            'Production Team': production || '',
            'Image & Video Generation Status': 'Not started',
            'Image & Video Generation Approval': 'Not yet approved',
            'Image & Video Generation Message': '',
            'Image & Video Generation Status - Finished Date': '',
            'Editor': editor || '',
            'Editing Status': 'Not started',
            'Editing Approval': 'Not yet approved',
            'Editing Message': '',
            'Editing Status - Finished Date': '',
            'Final Internal Approval': 'Not yet approved',
            'Client Approval': 'Not yet approved',
            'Broadcast Status': 'Not started'
        };

        const addedRow = await sheet.addRow(newRowData);
        res.json({ ok: true, task: addedRow.toObject() });
    } catch (err) {
        console.error(err);
        res.status(500).json({ error: err.message });
    }
});

app.post('/api/deleteTask', requireAuth, async (req, res) => {
    try {
        if (!req.isAdmin) {
            return res.status(403).json({ error: 'Only the admin can delete tasks.' });
        }
        const { videoTitle } = req.body;
        if (!videoTitle) return res.status(400).json({ error: 'Video Title is required.' });

        const sheet = await getSheet(MASTER_TAB);
        const row = await findRowByVideoTitle(sheet, videoTitle);
        if (!row) return res.status(404).json({ error: 'Video not found: ' + videoTitle });

        await row.delete();
        res.json({ ok: true });
    } catch (err) {
        console.error(err);
        res.status(500).json({ error: err.message });
    }
});

app.post('/api/startChanges', requireAuth, async (req, res) => {
    try {
        const { videoTitle, stageKey } = req.body;
        const statusFieldByStage = {
            script: 'Script & Storyboard Status',
            prod: 'Image & Video Generation Status',
            edit: 'Editing Status',
        };
        const roleFieldByStage = {
            script: 'Creative Team',
            prod: 'Production Team',
            edit: 'Editor',
        };
        const field = statusFieldByStage[stageKey];
        if (!field) return res.status(400).json({ error: 'Unknown stage: ' + stageKey });

        const sheet = await getSheet(MASTER_TAB);
        const row = await findRowByVideoTitle(sheet, videoTitle);
        if (!row) return res.status(404).json({ error: 'Video not found: ' + videoTitle });

        const assignedPerson = row.toObject()[roleFieldByStage[stageKey]];
        if (!req.isAdmin && assignedPerson !== req.userName) {
            return res.status(403).json({ error: `This task is assigned to ${assignedPerson || 'nobody'}, not you.` });
        }

        row.set(field, 'Changes in progress');
        await row.save();
        res.json({ ok: true });
    } catch (err) {
        console.error(err);
        res.status(500).json({ error: err.message });
    }
});

app.post('/api/updateStageLink', requireAuth, async (req, res) => {
    try {
        const { videoTitle, linkField, linkUrl } = req.body;
        const allowedFields = ['Creative Link', 'Production Link', 'Editor Link'];
        if (!allowedFields.includes(linkField)) {
            return res.status(400).json({ error: 'Invalid link field: ' + linkField });
        }

        const sheet = await getSheet(MASTER_TAB);
        const row = await findRowByVideoTitle(sheet, videoTitle);
        if (!row) return res.status(404).json({ error: 'Video not found: ' + videoTitle });

        const linkToRoleField = {
            'Creative Link': 'Creative Team',
            'Production Link': 'Production Team',
            'Editor Link': 'Editor'
        };
        const roleField = linkToRoleField[linkField];
        const assignedPerson = row.get(roleField);
        if (!req.isAdmin && assignedPerson !== req.userName) {
            return res.status(403).json({ error: `This task is assigned to ${assignedPerson || 'nobody'}, not you.` });
        }

        row.set(linkField, linkUrl || '');
        await row.save();
        res.json({ ok: true });
    } catch (err) {
        console.error(err);
        res.status(500).json({ error: err.message });
    }
});

app.post('/api/uploadAudio', requireAuth, async (req, res) => {
    try {
        const { audioBase64 } = req.body;
        if (!audioBase64) return res.status(400).json({ error: 'No audio data provided.' });

        const matches = audioBase64.match(/^data:([a-zA-Z0-9]+\/[a-zA-Z0-9-.+]+);base64,(.+)$/);
        if (!matches || matches.length !== 3) {
            return res.status(400).json({ error: 'Invalid base64 audio format.' });
        }

        const mimeType = matches[1];
        const buffer = Buffer.from(matches[2], 'base64');

        try {
            // 1. Fetch Access Token from JWT
            const tokenInfo = await auth.getAccessToken();
            const token = tokenInfo.token;

            // 2. Upload file to Google Drive
            const boundary = 'ccm_audio_boundary';
            const metadata = {
                name: `feedback-${Date.now()}.webm`,
                mimeType: mimeType
            };

            const multipartBody = Buffer.concat([
                Buffer.from(`--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify(metadata)}\r\n\r\n`),
                Buffer.from(`--${boundary}\r\nContent-Type: ${mimeType}\r\n\r\n`),
                buffer,
                Buffer.from(`\r\n--${boundary}--\r\n`)
            ]);

            const uploadRes = await fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,webViewLink,webContentLink', {
                method: 'POST',
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Content-Type': `multipart/related; boundary=${boundary}`
                },
                body: multipartBody
            });

            if (!uploadRes.ok) {
                const errText = await uploadRes.text();
                throw new Error(`Google Drive Upload failed: ${errText}`);
            }

            const fileData = await uploadRes.json();
            const fileId = fileData.id;

            // 3. Set reader permissions for anyone with the link
            const permRes = await fetch(`https://www.googleapis.com/drive/v3/files/${fileId}/permissions`, {
                method: 'POST',
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    role: 'reader',
                    type: 'anyone'
                })
            });

            if (!permRes.ok) {
                const errText = await permRes.text();
                throw new Error(`Google Drive Permissions failed: ${errText}`);
            }

            // Direct content download link
            const directUrl = fileData.webContentLink || fileData.webViewLink;
            res.json({ url: directUrl });
        } catch (driveErr) {
            console.warn('Google Drive API disabled or failed, falling back to local server storage:', driveErr.message);

            // FALLBACK: Store locally on the server!
            const dir = path.join(__dirname, 'public', 'uploads', 'audio');
            if (!fs.existsSync(dir)) {
                fs.mkdirSync(dir, { recursive: true });
            }
            const filename = `feedback-${Date.now()}.webm`;
            const filepath = path.join(dir, filename);
            fs.writeFileSync(filepath, buffer);

            res.json({ url: `/uploads/audio/${filename}` });
        }
    } catch (err) {
        console.error(err);
        res.status(500).json({ error: err.message });
    }
});

app.post('/api/addTeamMember', requireAuth, async (req, res) => {
    try {
        if (!req.isAdmin) {
            return res.status(403).json({ error: 'Only the admin can add team members.' });
        }
        const { name, team, email } = req.body;
        if (!name || !team || !email) {
            return res.status(400).json({ error: 'Name, Team/Role, and Email are required.' });
        }

        const trimmedName = name.trim();
        const trimmedEmail = email.trim().toLowerCase();
        const trimmedTeam = team.trim();

        if (ROSTER.some(p => p.name.toLowerCase() === trimmedName.toLowerCase())) {
            return res.status(400).json({ error: `Member with name "${trimmedName}" already exists.` });
        }
        if (TEAM_ACCOUNTS[trimmedEmail]) {
            return res.status(400).json({ error: `Account with email "${trimmedEmail}" already exists.` });
        }

        // Add to roster
        ROSTER.push({ name: trimmedName, team: trimmedTeam });
        fs.writeFileSync(ROSTER_FILE, JSON.stringify(ROSTER, null, 2), 'utf8');

        // Add to team accounts
        TEAM_ACCOUNTS[trimmedEmail] = { name: trimmedName };
        fs.writeFileSync(TEAM_ACCOUNTS_FILE, JSON.stringify(TEAM_ACCOUNTS, null, 2), 'utf8');

        refreshTeamEmails();
        res.json({ ok: true, roster: ROSTER, teamAccounts: TEAM_ACCOUNTS });
    } catch (err) {
        console.error(err);
        res.status(500).json({ error: err.message });
    }
});

app.post('/api/resetUserPassword', requireAuth, async (req, res) => {
    try {
        if (!req.isAdmin) {
            return res.status(403).json({ error: 'Only the admin can reset passwords.' });
        }
        const { email } = req.body;
        if (!email) return res.status(400).json({ error: 'Email is required.' });

        const emailLc = email.toLowerCase().trim();
        if (userPasswords[emailLc]) {
            delete userPasswords[emailLc];
            try {
                fs.writeFileSync(PASSWORDS_FILE, JSON.stringify(userPasswords, null, 2), 'utf8');
            } catch (e) {
                console.error('Error saving passwords.json:', e);
            }
        }
        res.json({ ok: true });
    } catch (err) {
        console.error(err);
        res.status(500).json({ error: err.message });
    }
});

app.listen(process.env.PORT, () => console.log(`Server running on http://localhost:${process.env.PORT}`));
