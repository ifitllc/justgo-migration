const fs = require('fs/promises');
const path = require('path');
const XLSX = require('xlsx');

const TOURNAMENTS_DIR = path.join(__dirname, 'tournaments');
const TEMPLATE_PATH = path.join(__dirname, 'USATT tournament results template.xls');
const OUTPUT_FILENAME = 'justgo-results.xlsx';

const MATCH_SHEET = 'Match Results';
const ESTIMATED_SHEET = 'Estimated Ratings';

const MATCH_HEADERS = [
	'Winner Name',
	'Winner Membership#',
	'Loser Name',
	'Loser Membership#',
	'Scores',
	'Event'
];

const ESTIMATED_HEADERS = ['Name', 'Membership#', 'Est Rating '];

const normalizeString = (value) => (typeof value === 'string' ? value.trim() : value ?? '');

const normalizeMembershipNumber = (value) => {
	const str = normalizeString(value);
	if (!str) return '';
	const numeric = Number(str);
	if (!Number.isNaN(numeric)) return String(Math.trunc(numeric));
	return str;
};

const formatName = (first, last) => `${normalizeString(first)} ${normalizeString(last)}`.trim();

const resolveTournamentFolder = () => {
	const arg = normalizeString(process.argv[2]);
	if (!arg) return '';
	return arg;
};

const findLatestFile = async (dirPath, token) => {
	const entries = await fs.readdir(dirPath, { withFileTypes: true });
	const candidates = entries
		.filter((entry) => entry.isFile() && entry.name.includes(token))
		.map((entry) => entry.name);

	if (!candidates.length) return null;

	const stats = await Promise.all(
		candidates.map(async (name) => ({
			name,
			time: (await fs.stat(path.join(dirPath, name))).mtimeMs
		}))
	);

	return stats.sort((a, b) => b.time - a.time)[0].name;
};

const readCsvAsJson = async (filePath) => {
	const csv = await fs.readFile(filePath, 'utf8');
	const workbook = XLSX.read(csv, { type: 'string' });
	const sheet = workbook.Sheets[workbook.SheetNames[0]];
	return XLSX.utils.sheet_to_json(sheet, { defval: '' });
};

const loadInputs = async (folder) => {
	const folderPath = path.isAbsolute(folder) ? folder : path.join(TOURNAMENTS_DIR, folder);
	const matchFile = await findLatestFile(folderPath, 'match-results');
	const playersFile = await findLatestFile(folderPath, 'players');
	const membershipsFile = await findLatestFile(folderPath, 'usatt-memberships');

	if (!matchFile || !playersFile || !membershipsFile) {
		throw new Error('Could not locate match-results, players, and usatt-memberships CSV files.');
	}

	const [matchResults, players, memberships] = await Promise.all([
		readCsvAsJson(path.join(folderPath, matchFile)),
		readCsvAsJson(path.join(folderPath, playersFile)),
		readCsvAsJson(path.join(folderPath, membershipsFile))
	]);

	return { folderPath, matchResults, players, memberships };
};

const buildMembershipMap = (memberships) => {
	const byNumber = new Map();

	memberships.forEach((row) => {
		const membershipNumber = normalizeMembershipNumber(row['Membership#']);
		if (!membershipNumber) return;
		byNumber.set(membershipNumber, {
			membershipNumber,
			firstName: normalizeString(row.FirstName),
			lastName: normalizeString(row.LastName),
			estRating: normalizeString(row.EstRating)
		});
	});

	return byNumber;
};

const buildPlayerMaps = (players, membershipMap) => {
	const membershipLookup = new Map();
	const membershipToPlayer = new Map();
	const playerById = new Map();

	players.forEach((player) => {
		const playerId = normalizeString(player.id);
		const memberId = normalizeMembershipNumber(player.memberId);
		const membershipNumber = memberId
			? membershipMap.get(memberId)?.membershipNumber || memberId
			: '';

		if (playerId) playerById.set(playerId, player);
		if (membershipNumber) {
			membershipLookup.set(playerId, membershipNumber);
			membershipLookup.set(memberId, membershipNumber);
			membershipToPlayer.set(membershipNumber, player);
		}
	});

	return { membershipLookup, membershipToPlayer, playerById };
};

const resolveMembership = (key, membershipLookup) => {
	const normalized = normalizeMembershipNumber(key);
	if (!normalized) return '';
	return membershipLookup.get(normalized) || normalized;
};

const resolveName = (membershipNumber, membershipMap, membershipToPlayer, playerById, fallbackKey) => {
	if (membershipNumber) {
		const membership = membershipMap.get(membershipNumber);
		if (membership) return formatName(membership.firstName, membership.lastName);

		const player = membershipToPlayer.get(membershipNumber);
		if (player) return formatName(player.firstName, player.lastName);
	}

	if (fallbackKey) {
		const player = playerById.get(normalizeString(fallbackKey));
		if (player) return formatName(player.firstName, player.lastName);
	}

	return '';
};

const mapMatchResults = (matchResults, helpers) => {
	const { membershipLookup, membershipMap, membershipToPlayer, playerById } = helpers;

	return matchResults
		.map((row) => {
			const winnerKey = row.MemNum_W;
			const loserKey = row.MemNum_L;

			const winnerMembership = resolveMembership(winnerKey, membershipLookup);
			const loserMembership = resolveMembership(loserKey, membershipLookup);

			const winnerName = resolveName(winnerMembership, membershipMap, membershipToPlayer, playerById, winnerKey);
			const loserName = resolveName(loserMembership, membershipMap, membershipToPlayer, playerById, loserKey);

			const scores = normalizeString(row.Score);
			const event = normalizeString(row.Division);

			if (!winnerMembership && !loserMembership && !scores && !event) return null;

			return [winnerName, winnerMembership, loserName, loserMembership, scores, event];
		})
		.filter(Boolean);
};

const mapEstimatedRatings = (players, helpers) => {
	const { membershipLookup, membershipMap } = helpers;

	return players.map((player) => {
		const membershipNumber = resolveMembership(player.memberId || player.id, membershipLookup);
		const estRating = membershipNumber ? membershipMap.get(membershipNumber)?.estRating || '' : '';
		const name = formatName(player.firstName, player.lastName);
		return [name, membershipNumber, estRating];
	});
};

const writeWorkbook = (matchRows, estimatedRows, outputPath) => {
	const workbook = XLSX.readFile(TEMPLATE_PATH);
	const matchSheet = XLSX.utils.aoa_to_sheet([MATCH_HEADERS, ...matchRows]);
	const estimatedSheet = XLSX.utils.aoa_to_sheet([ESTIMATED_HEADERS, ...estimatedRows]);

	const replaceSheet = (name, sheet) => {
		const index = workbook.SheetNames.indexOf(name);
		if (index === -1) {
			XLSX.utils.book_append_sheet(workbook, sheet, name);
			return;
		}
		workbook.Sheets[name] = sheet;
		workbook.SheetNames[index] = name;
	};

	replaceSheet(MATCH_SHEET, matchSheet);
	replaceSheet(ESTIMATED_SHEET, estimatedSheet);

	XLSX.writeFile(workbook, outputPath);
};

const ensureFolder = async (folderPath) => {
	await fs.mkdir(folderPath, { recursive: true });
};

const run = async () => {
	const folderArg = resolveTournamentFolder();
	if (!folderArg) {
		throw new Error('Provide a tournament folder name or path (e.g., 202512).');
	}

	const { folderPath, matchResults, players, memberships } = await loadInputs(folderArg);
	const membershipMap = buildMembershipMap(memberships);
	const helperMaps = buildPlayerMaps(players, membershipMap);
	const matchRows = mapMatchResults(matchResults, {
		membershipLookup: helperMaps.membershipLookup,
		membershipMap,
		membershipToPlayer: helperMaps.membershipToPlayer,
		playerById: helperMaps.playerById
	});
	const estimatedRows = mapEstimatedRatings(players, {
		membershipLookup: helperMaps.membershipLookup,
		membershipMap
	});

	const outputPath = path.join(folderPath, OUTPUT_FILENAME);
	await ensureFolder(folderPath);
	writeWorkbook(matchRows, estimatedRows, outputPath);

	console.log(`Wrote Match Results rows: ${matchRows.length}`);
	console.log(`Wrote Estimated Ratings rows: ${estimatedRows.length}`);
	console.log(`Output: ${outputPath}`);
};

run().catch((error) => {
	console.error(error.message || error);
	process.exitCode = 1;
});
