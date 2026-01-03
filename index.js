const fs = require('fs/promises');
const path = require('path');
const XLSX = require('xlsx');

const TOURNAMENTS_DIR = path.join(__dirname, 'tournaments');
const OUTPUT_FILENAME = 'justgo-import.xlsx';
const TEMPLATE_FILENAME = 'USATT tournament results template.xls';

const HEADER = [
	'Firstname*',
	'Lastname*',
	'EmailAddress*',
	'DOB*',
	'Username*',
	'Gender',
	'Title',
	'Address1',
	'Address2',
	'Town',
	'County',
	'PostCode',
	'Country',
	'Mobile Telephone',
	'Home Telephone',
	'Emergency Contact First Name',
	'Emergency Contact Surname',
	'Emergency Contact Relationship',
	'Emergency Contact Number',
	'Emergency Contact Email Address',
	'Parent FirstName',
	'Parent Surname',
	'Parent EmailAddress'
];

const REQUIRED_FIELDS = ['firstName', 'lastName', 'email', 'dob'];

const formatDateSegment = (value) => String(value).padStart(2, '0');

const formatFolderDate = (date) => {
	const year = date.getFullYear();
	const month = formatDateSegment(date.getMonth() + 1);
	const day = formatDateSegment(date.getDate());
	return `${year}-${month}-${day}`;
};

const normalizeString = (value) => (typeof value === 'string' ? value.trim() : '');

const resolveTournamentFolder = () => {
	const arg = normalizeString(process.argv[2]);
	if (!arg) return formatFolderDate(new Date());

	if (!/^\d{6}$/.test(arg)) {
		throw new Error('Tournament parameter must be yyyymm');
	}

	return arg;
};

const normalizeDob = (value) => {
	const raw = normalizeString(value);
	if (!raw) return raw;

	const parsed = new Date(raw);
	if (Number.isNaN(parsed.getTime())) return raw;

	const month = formatDateSegment(parsed.getMonth() + 1);
	const day = formatDateSegment(parsed.getDate());
	const year = parsed.getFullYear();
	return `${month}/${day}/${year}`;
};

const buildUsername = (player) => {
	const email = normalizeString(player.email);
	if (email) return email;
	const first = normalizeString(player.firstName);
	const last = normalizeString(player.lastName);
	return [first, last].filter(Boolean).join('.').toLowerCase();
};

const mapPlayerToRow = (player) => {
	const missing = REQUIRED_FIELDS.filter((field) => !normalizeString(player[field]));

	return {
		row: [
			normalizeString(player.firstName),
			normalizeString(player.lastName),
			normalizeString(player.email),
			normalizeDob(player.dob),
			buildUsername(player),
			normalizeString(player.gender),
			'',
			normalizeString(player.address),
			'',
			normalizeString(player.city),
			normalizeString(player.state),
			'',
			'',
			normalizeString(player.phone),
			'',
			'',
			'',
			'',
			'',
			'',
			'',
			'',
			''
		],
		missing
	};
};

const loadPlayers = async (playersPath) => {
	const raw = await fs.readFile(playersPath, 'utf8');
	const players = JSON.parse(raw);

	if (!Array.isArray(players)) {
		throw new Error('players.json must contain an array');
	}

	return players;
};

const writeWorkbook = async (rows, outputPath) => {
	const worksheet = XLSX.utils.aoa_to_sheet([HEADER, ...rows]);
	const workbook = XLSX.utils.book_new();
	XLSX.utils.book_append_sheet(workbook, worksheet, 'Members');
	XLSX.writeFile(workbook, outputPath);
};

// CSV generation removed: no longer producing justgo-players.csv

const writeEstimatedRatingsToTemplate = async (players, templatePath, outputPath) => {
	const sheetName = 'Estimated Ratings';

	// Read the template workbook (may be .xls)
	const workbook = XLSX.readFile(templatePath, { cellDates: true });

	// Build the data array: header + rows
	const header = ['Name', 'Membership#', 'Est Rating'];
	const data = [header];

	players.forEach((p) => {
		const name = [p.firstName, p.lastName].map(normalizeString).filter(Boolean).join(' ');
		const membership = normalizeString(p.usattId || '');
		const rating = p.rating == null ? '' : String(p.rating);
		data.push([name, membership, rating]);
	});

	const sheet = XLSX.utils.aoa_to_sheet(data);
	workbook.Sheets[sheetName] = sheet;
	if (!workbook.SheetNames.includes(sheetName)) workbook.SheetNames.push(sheetName);

	XLSX.writeFile(workbook, outputPath);
};

const ensureOutputDir = async (dirPath) => {
	await fs.mkdir(dirPath, { recursive: true });
};

const run = async () => {
	const tournamentFolder = resolveTournamentFolder();
	const playersPath = path.join(TOURNAMENTS_DIR, tournamentFolder, 'players.json');
	const players = await loadPlayers(playersPath);
	const rows = [];
	const missingReport = [];

	players.forEach((player, index) => {
		const { row, missing } = mapPlayerToRow(player);
		rows.push(row);
		if (missing.length) {
			missingReport.push({
				index,
				name: `${normalizeString(player.firstName)} ${normalizeString(player.lastName)}`.trim(),
				missing
			});
		}
	});

	const outputDir = path.join(TOURNAMENTS_DIR, tournamentFolder);
	await ensureOutputDir(outputDir);

	const outputPath = path.join(outputDir, OUTPUT_FILENAME);
	try {
		await writeWorkbook(rows, outputPath);
		console.log(`Wrote ${rows.length} player(s) to ${outputPath}`);
	} catch (err) {
		console.warn(`Failed to write workbook: ${err.message}`);
	}

	// Create HCTT results workbook based on the USATT template and populate
	// the "Estimated Ratings" sheet with Name, Membership#, Est Rating.
	try {
		const templatePath = path.join(__dirname, TEMPLATE_FILENAME);
		let dateStr = tournamentFolder;
		if (/^\d{6}$/.test(tournamentFolder)) {
			dateStr = `${tournamentFolder.slice(0,4)}-${tournamentFolder.slice(4,6)}-01`;
		}
		const resultsFilename = `HCTT ${dateStr} Results.xlsx`;
		const resultsPath = path.join(outputDir, resultsFilename);
		await writeEstimatedRatingsToTemplate(players, templatePath, resultsPath);
		console.log(`Wrote Estimated Ratings to ${resultsPath}`);
	} catch (err) {
		console.warn(`Failed to write HCTT results workbook: ${err.message}`);
	}

	if (missingReport.length) {
		console.warn('Players with missing required fields:');
		missingReport.forEach((entry) => {
			const displayName = entry.name || '(no name)';
			console.warn(`- #${entry.index + 1} ${displayName} -> missing: ${entry.missing.join(', ')}`);
		});
	}
};

run().catch((error) => {
	console.error(error);
	process.exitCode = 1;
});
