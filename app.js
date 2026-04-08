(function() {
	'use strict';

	var ONE_DRIVE_CONFIG = {
		clientId: '',
		authority: 'https://login.microsoftonline.com/common',
		redirectUri: window.location.origin + window.location.pathname,
		scopes: ['Files.Read']
	};

	var state = {
		entries: [],
		countsByDay: {},
		selectedDate: startOfDay(new Date()),
		currentMonth: firstDayOfMonth(new Date()),
		activeTab: 'month',
		theme: 'desktop'
	};

	var authClient = null;

	var els = {
		body: document.body,
		fileBtn: document.getElementById('load-file-btn'),
		oneDriveBtn: document.getElementById('load-onedrive-btn'),
		fileInput: document.getElementById('local-file-input'),
		sourceStatus: document.getElementById('source-status'),
		summary: document.getElementById('summary'),
		summaryTotal: document.getElementById('summary-total'),
		summaryDays: document.getElementById('summary-days'),
		summaryBooks: document.getElementById('summary-books'),
		mainPanel: document.getElementById('main-panel'),
		tabMonth: document.getElementById('tab-month'),
		tabDay: document.getElementById('tab-day'),
		monthView: document.getElementById('month-view'),
		dayView: document.getElementById('day-view'),
		monthPrev: document.getElementById('month-prev'),
		monthNext: document.getElementById('month-next'),
		monthLabel: document.getElementById('month-label'),
		monthGrid: document.getElementById('month-grid'),
		dayPrev: document.getElementById('day-prev'),
		dayNext: document.getElementById('day-next'),
		dayLabel: document.getElementById('day-label'),
		dayEntries: document.getElementById('day-entries'),
		desktopThemeBtn: document.getElementById('desktop-theme-btn'),
		kindleThemeBtn: document.getElementById('kindle-theme-btn')
	};

	init();

	function init() {
		loadThemePreference();
		bindEvents();
		renderSourceStatus('Choose a source to load Kindle notes.', 'info');
		renderAll();
	}

	function bindEvents() {
		els.fileBtn.addEventListener('click', function() {
			els.fileInput.click();
		});

		els.fileInput.addEventListener('change', onLocalFileChange);

		els.oneDriveBtn.addEventListener('click', function() {
			onOneDriveImport();
		});

		els.tabMonth.addEventListener('click', function() {
			switchTab('month');
		});

		els.tabDay.addEventListener('click', function() {
			switchTab('day');
		});

		els.monthPrev.addEventListener('click', function() {
			state.currentMonth = new Date(state.currentMonth.getFullYear(), state.currentMonth.getMonth() - 1, 1);
			renderMonthView();
		});

		els.monthNext.addEventListener('click', function() {
			state.currentMonth = new Date(state.currentMonth.getFullYear(), state.currentMonth.getMonth() + 1, 1);
			renderMonthView();
		});

		els.dayPrev.addEventListener('click', function() {
			shiftDay(-1);
		});

		els.dayNext.addEventListener('click', function() {
			shiftDay(1);
		});

		els.desktopThemeBtn.addEventListener('click', function() {
			setTheme('desktop');
		});

		els.kindleThemeBtn.addEventListener('click', function() {
			setTheme('kindle');
		});
	}

	function onLocalFileChange(event) {
		var file = event.target.files && event.target.files[0];

		if (!file) {
			return;
		}

		var reader = new FileReader();

		reader.onload = function(loadEvent) {
			var text = String(loadEvent.target.result || '');
			loadEntries(text, 'Imported from local file: ' + file.name + '.');
		};

		reader.onerror = function() {
			renderSourceStatus('Unable to read that file. Please try another export.', 'warn');
		};

		reader.readAsText(file);
		event.target.value = '';
	}

	function onOneDriveImport() {
		if (!window.msal) {
			renderSourceStatus('MSAL library could not be loaded. Check network access and try again.', 'warn');
			return;
		}

		if (!ONE_DRIVE_CONFIG.clientId) {
			renderSourceStatus('OneDrive is not configured yet. Add your Azure app clientId in app.js.', 'warn');
			return;
		}

		renderSourceStatus('Connecting to OneDrive and searching for My Clippings.txt...', 'info');

		ensureAuthClient()
			.then(function(client) {
				return acquireGraphToken(client);
			})
			.then(function(token) {
				return findClippingsFile(token).then(function(fileMeta) {
					return downloadDriveItemText(token, fileMeta.id).then(function(text) {
						return {
							fileMeta: fileMeta,
							text: text
						};
					});
				});
			})
			.then(function(result) {
				loadEntries(result.text, 'Imported from OneDrive file: ' + result.fileMeta.name + '.');
			})
			.catch(function(error) {
				var message = error && error.message ? error.message : 'OneDrive import failed.';
				renderSourceStatus(message, 'warn');
			});
	}

	function ensureAuthClient() {
		if (authClient) {
			return Promise.resolve(authClient);
		}

		try {
			authClient = new window.msal.PublicClientApplication({
				auth: {
					clientId: ONE_DRIVE_CONFIG.clientId,
					authority: ONE_DRIVE_CONFIG.authority,
					redirectUri: ONE_DRIVE_CONFIG.redirectUri
				},
				cache: {
					cacheLocation: 'sessionStorage'
				}
			});
			return Promise.resolve(authClient);
		} catch (error) {
			return Promise.reject(error);
		}
	}

	function acquireGraphToken(client) {
		var accounts = client.getAllAccounts();
		var loginPromise;

		if (accounts.length > 0) {
			loginPromise = Promise.resolve({ account: accounts[0] });
		} else {
			loginPromise = client.loginPopup({ scopes: ONE_DRIVE_CONFIG.scopes });
		}

		return loginPromise.then(function(loginResult) {
			var tokenRequest = {
				account: loginResult.account,
				scopes: ONE_DRIVE_CONFIG.scopes
			};

			return client.acquireTokenSilent(tokenRequest).catch(function() {
				return client.acquireTokenPopup(tokenRequest);
			});
		}).then(function(tokenResult) {
			return tokenResult.accessToken;
		});
	}

	function findClippingsFile(accessToken) {
		var url = 'https://graph.microsoft.com/v1.0/me/drive/root/search(q=\'My Clippings.txt\')?$select=id,name,lastModifiedDateTime';

		return fetch(url, {
			headers: {
				Authorization: 'Bearer ' + accessToken
			}
		}).then(function(response) {
			if (!response.ok) {
				throw new Error('Could not search OneDrive for My Clippings.txt.');
			}
			return response.json();
		}).then(function(data) {
			if (!data || !data.value || !data.value.length) {
				throw new Error('My Clippings.txt was not found in this OneDrive.');
			}

			data.value.sort(function(a, b) {
				var aTime = new Date(a.lastModifiedDateTime || 0).getTime();
				var bTime = new Date(b.lastModifiedDateTime || 0).getTime();
				return bTime - aTime;
			});

			return data.value[0];
		});
	}

	function downloadDriveItemText(accessToken, itemId) {
		var url = 'https://graph.microsoft.com/v1.0/me/drive/items/' + encodeURIComponent(itemId) + '/content';

		return fetch(url, {
			headers: {
				Authorization: 'Bearer ' + accessToken
			}
		}).then(function(response) {
			if (!response.ok) {
				throw new Error('OneDrive file was found, but it could not be downloaded.');
			}
			return response.text();
		});
	}

	function loadEntries(text, sourceLabel) {
		state.entries = parseClippings(text);
		state.countsByDay = countByDay(state.entries);
		chooseInitialDate();
		state.currentMonth = firstDayOfMonth(state.selectedDate);

		if (!state.entries.length) {
			renderSourceStatus(sourceLabel + ' No entries were parsed. Verify this is a full Kindle clippings export.', 'warn');
		} else {
			renderSourceStatus(sourceLabel + ' Parsed ' + state.entries.length + ' entries.', 'success');
		}

		els.mainPanel.classList.remove('hidden');
		renderSummary();
		renderAll();
	}

	function chooseInitialDate() {
		var todayKey = toDateKey(new Date());

		if (state.countsByDay[todayKey]) {
			state.selectedDate = startOfDay(new Date());
			return;
		}

		if (state.entries.length) {
			state.selectedDate = startOfDay(state.entries[state.entries.length - 1].date);
			return;
		}

		state.selectedDate = startOfDay(new Date());
	}

	function parseClippings(rawText) {
		var blocks = String(rawText || '').split('==========');
		var parsed = [];
		var i;

		for (i = 0; i < blocks.length; i += 1) {
			var block = blocks[i].replace(/^\s+|\s+$/g, '');
			if (!block) {
				continue;
			}

			var lines = block.split(/\r?\n/);
			if (lines.length < 2) {
				continue;
			}

			var titleLine = normalizeWhitespace(lines[0]);
			var metadataLine = normalizeWhitespace(lines[1]);
			var addedOn = extractAddedOn(metadataLine);
			var parsedDate = parseKindleDate(addedOn);

			if (!parsedDate) {
				continue;
			}

			var details = splitTitleAndAuthor(titleLine);
			var content = lines.slice(2).join('\n').replace(/^\s+|\s+$/g, '');

			parsed.push({
				id: 'entry-' + i + '-' + parsedDate.getTime(),
				title: details.title,
				author: details.author,
				type: detectClipType(metadataLine),
				date: parsedDate,
				dateKey: toDateKey(parsedDate),
				content: content || '(No note text)'
			});
		}

		parsed.sort(function(a, b) {
			return a.date.getTime() - b.date.getTime();
		});

		return parsed;
	}

	function extractAddedOn(metadataLine) {
		var match = metadataLine.match(/Added on (.+)$/i);
		return match && match[1] ? match[1] : '';
	}

	function splitTitleAndAuthor(titleLine) {
		var match = titleLine.match(/^(.*)\(([^()]*)\)\s*$/);

		if (!match) {
			return {
				title: titleLine,
				author: 'Unknown author'
			};
		}

		return {
			title: match[1].replace(/\s+$/g, ''),
			author: match[2].replace(/^\s+|\s+$/g, '') || 'Unknown author'
		};
	}

	function detectClipType(metadataLine) {
		var match = metadataLine.match(/Your\s+([A-Za-z]+)/i);

		if (!match || !match[1]) {
			return 'Clip';
		}

		var lower = match[1].toLowerCase();
		return lower.charAt(0).toUpperCase() + lower.slice(1);
	}

	function parseKindleDate(value) {
		if (!value) {
			return null;
		}

		var normalized = normalizeWhitespace(String(value))
			.replace(/\sat\s/i, ' ')
			.replace(/^([A-Za-z]+),\s+/, '$1 ');

		var attempts = [
			String(value),
			normalized,
			normalized.replace(/^[A-Za-z]+\s+/, '')
		];

		var i;
		for (i = 0; i < attempts.length; i += 1) {
			var candidate = new Date(attempts[i]);
			if (!isNaN(candidate.getTime())) {
				return candidate;
			}
		}

		return null;
	}

	function countByDay(entries) {
		var map = {};
		var i;

		for (i = 0; i < entries.length; i += 1) {
			var key = entries[i].dateKey;
			map[key] = (map[key] || 0) + 1;
		}

		return map;
	}

	function renderAll() {
		renderTabs();
		renderMonthView();
		renderDayView();
	}

	function renderTabs() {
		var monthActive = state.activeTab === 'month';

		els.tabMonth.classList.toggle('active', monthActive);
		els.tabDay.classList.toggle('active', !monthActive);
		els.tabMonth.setAttribute('aria-selected', monthActive ? 'true' : 'false');
		els.tabDay.setAttribute('aria-selected', monthActive ? 'false' : 'true');
		els.monthView.classList.toggle('active', monthActive);
		els.dayView.classList.toggle('active', !monthActive);
	}

	function switchTab(tabName) {
		state.activeTab = tabName;
		renderTabs();
	}

	function renderMonthView() {
		var year = state.currentMonth.getFullYear();
		var month = state.currentMonth.getMonth();
		var firstDay = new Date(year, month, 1);
		var startWeekday = firstDay.getDay();
		var totalDays = new Date(year, month + 1, 0).getDate();
		var weekdays = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
		var slot;

		els.monthLabel.textContent = formatMonthLabel(state.currentMonth);
		els.monthGrid.innerHTML = '';

		for (slot = 0; slot < weekdays.length; slot += 1) {
			var weekdayCell = document.createElement('div');
			weekdayCell.className = 'weekday';
			weekdayCell.textContent = weekdays[slot];
			els.monthGrid.appendChild(weekdayCell);
		}

		for (slot = 0; slot < 42; slot += 1) {
			var dayNumber = slot - startWeekday + 1;

			if (dayNumber < 1 || dayNumber > totalDays) {
				var emptyCell = document.createElement('button');
				emptyCell.className = 'day-cell empty';
				emptyCell.type = 'button';
				emptyCell.disabled = true;
				els.monthGrid.appendChild(emptyCell);
				continue;
			}

			appendDayCell(year, month, dayNumber);
		}
	}

	function appendDayCell(year, month, dayNumber) {
		var dayDate = new Date(year, month, dayNumber);
		var key = toDateKey(dayDate);
		var count = state.countsByDay[key] || 0;
		var button = document.createElement('button');
		var dayText = document.createElement('span');
		var countText = document.createElement('span');

		button.type = 'button';
		button.className = 'day-cell';
		if (count > 0) {
			button.className += ' has-count';
		}

		if (key === toDateKey(state.selectedDate)) {
			button.className += ' selected';
		}

		dayText.className = 'day-number';
		dayText.textContent = String(dayNumber);

		countText.className = 'count';
		countText.textContent = count === 1 ? '1 entry' : count + ' entries';

		button.appendChild(dayText);
		button.appendChild(countText);

		button.addEventListener('click', function() {
			state.selectedDate = startOfDay(dayDate);
			state.currentMonth = firstDayOfMonth(dayDate);
			state.activeTab = 'day';
			renderAll();
		});

		els.monthGrid.appendChild(button);
	}

	function renderDayView() {
		var date = state.selectedDate;
		var entries = entriesForDate(date);
		var i;

		els.dayLabel.textContent = formatDayLabel(date) + ' (' + entries.length + ')';
		els.dayEntries.innerHTML = '';

		if (!entries.length) {
			var empty = document.createElement('p');
			empty.className = 'empty-note';
			empty.textContent = 'No notes or highlights found for this day.';
			els.dayEntries.appendChild(empty);
			return;
		}

		for (i = 0; i < entries.length; i += 1) {
			var entryCard = buildEntryCard(entries[i]);
			entryCard.style.animationDelay = (i * 45) + 'ms';
			els.dayEntries.appendChild(entryCard);
		}
	}

	function buildEntryCard(entry) {
		var card = document.createElement('article');
		var head = document.createElement('div');
		var tag = document.createElement('span');
		var time = document.createElement('span');
		var title = document.createElement('h3');
		var author = document.createElement('p');
		var text = document.createElement('p');

		card.className = 'clip';
		head.className = 'clip-head';
		tag.className = 'tag';
		tag.textContent = entry.type;
		time.textContent = formatTime(entry.date);
		head.appendChild(tag);
		head.appendChild(time);

		title.textContent = entry.title;
		author.className = 'author';
		author.textContent = entry.author;
		text.textContent = entry.content;

		card.appendChild(head);
		card.appendChild(title);
		card.appendChild(author);
		card.appendChild(text);

		return card;
	}

	function renderSummary() {
		var total = state.entries.length;
		var days = Object.keys(state.countsByDay).length;
		var booksMap = {};
		var i;

		for (i = 0; i < state.entries.length; i += 1) {
			booksMap[state.entries[i].title] = true;
		}

		els.summaryTotal.textContent = String(total);
		els.summaryDays.textContent = String(days);
		els.summaryBooks.textContent = String(Object.keys(booksMap).length);
		els.summary.classList.toggle('hidden', total === 0);
	}

	function shiftDay(offset) {
		state.selectedDate = new Date(
			state.selectedDate.getFullYear(),
			state.selectedDate.getMonth(),
			state.selectedDate.getDate() + offset
		);
		state.currentMonth = firstDayOfMonth(state.selectedDate);
		renderAll();
	}

	function entriesForDate(date) {
		var key = toDateKey(date);
		var filtered = [];
		var i;

		for (i = 0; i < state.entries.length; i += 1) {
			if (state.entries[i].dateKey === key) {
				filtered.push(state.entries[i]);
			}
		}

		filtered.sort(function(a, b) {
			return a.date.getTime() - b.date.getTime();
		});

		return filtered;
	}

	function renderSourceStatus(message, type) {
		els.sourceStatus.textContent = message;
		els.sourceStatus.className = 'status ' + type;
	}

	function setTheme(theme) {
		state.theme = theme === 'kindle' ? 'kindle' : 'desktop';
		applyTheme();

		try {
			window.localStorage.setItem('kindleNotesTheme', state.theme);
		} catch (error) {
			// Ignore storage errors and keep the current theme.
		}
	}

	function loadThemePreference() {
		try {
			var storedTheme = window.localStorage.getItem('kindleNotesTheme');
			state.theme = storedTheme === 'kindle' ? 'kindle' : 'desktop';
		} catch (error) {
			state.theme = 'desktop';
		}
		applyTheme();
	}

	function applyTheme() {
		els.body.classList.toggle('theme-kindle', state.theme === 'kindle');
		els.body.classList.toggle('theme-desktop', state.theme !== 'kindle');
		els.desktopThemeBtn.classList.toggle('active', state.theme === 'desktop');
		els.kindleThemeBtn.classList.toggle('active', state.theme === 'kindle');
	}

	function toDateKey(date) {
		return [
			date.getFullYear(),
			pad(date.getMonth() + 1),
			pad(date.getDate())
		].join('-');
	}

	function formatMonthLabel(date) {
		return date.toLocaleDateString(undefined, {
			month: 'long',
			year: 'numeric'
		});
	}

	function formatDayLabel(date) {
		return date.toLocaleDateString(undefined, {
			weekday: 'long',
			month: 'long',
			day: 'numeric',
			year: 'numeric'
		});
	}

	function formatTime(date) {
		return date.toLocaleTimeString(undefined, {
			hour: '2-digit',
			minute: '2-digit'
		});
	}

	function firstDayOfMonth(date) {
		return new Date(date.getFullYear(), date.getMonth(), 1);
	}

	function startOfDay(date) {
		return new Date(date.getFullYear(), date.getMonth(), date.getDate());
	}

	function pad(value) {
		return value < 10 ? '0' + value : String(value);
	}

	function normalizeWhitespace(value) {
		return String(value || '').replace(/\s+/g, ' ').replace(/^\s+|\s+$/g, '');
	}

})();