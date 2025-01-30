function onOpen() {
	const ui = DocumentApp.getUi();
	ui.createMenu('Custom Tools')
		.addItem('Insert stamps', 'insertStamps')
		.addToUi();
}

// Run nowStamp and mmddStamp for all tabs
function insertStamps() {

	nowStampedTabs = alterAllTabs(nowStamp);
	mmddStampedTabs = alterAllTabs(mmddStamp);

	let nowStampedString = 'nowStamped instances:'
	let mmddStampedString = 'mmddStamped instances:'
	Object.keys(nowStampedTabs).forEach(key => {
		if (mmddStampedTabs[key] > 0) {
			nowStampedString += `\n${nowStampedTabs[key]} in ${key}`;
		}
	})
	Object.keys(mmddStampedTabs).forEach(key => {
		if (mmddStampedTabs[key] > 0) {
			mmddStampedString += `\n${mmddStampedTabs[key]} in ${key}`;
		}
	})

	DocumentApp.getUi().alert(`${nowStampedString}\n${mmddStampedString}`);
}

// Perform funcToUse on all tabs
function alterAllTabs(funcToUse) {

	const tabs = DocumentApp.getActiveDocument().getTabs();

	// Keep track of number of alterations
	let alteredTabs = {};

	tabs.forEach(tab => {
		alteredChildTabs = alterChildTabs(tab, funcToUse, alteredTabs);
		Object.assign(alteredTabs, alteredChildTabs);
	});

	return alteredTabs;
}

// Perform funcToUse on parentTab and recursively call for child tabs
function alterChildTabs(parentTab, funcToUse, alteredTabs) {

	// Perform funcToUse on parentTab
	alteredParentTabs = funcToUse(parentTab.getId());
	alteredTabs[parentTab.getTitle()] = alteredParentTabs;

	// Call self for child tabs
	parentTab.getChildTabs().forEach(tab => {
		if (tab.getChildTabs().length > 0) {
			alteredChildTabs = alterChildTabs(tab, funcToUse, alteredTabs);
			Object.assign(alteredTabs, alteredChildTabs);
		} else {
			alteredTabs[tab.getTitle()] = funcToUse(tab.getId());
		}
	})

	return alteredTabs;
}

// Replace 'nowStamp' with current timestamp and return # matches replaces
function nowStamp(tabId) {

	const tab = DocumentApp.getActiveDocument().getTab(tabId);
	const body = tab.asDocumentTab().getBody();

	const now = new Date();
	const formattedDateTime = Utilities.formatDate(now, Session.getScriptTimeZone(), "MM-dd-yyyy-HH-mm");

	const text = body.getText();
	const matches = text.match(/\bnowStamp\b/g);
	if (matches) {
		// Loop through matches and replace them
		matches.forEach(match => {
			// Replace the match in the document
			body.replaceText(match, formattedDateTime);
		});
	}
	if (matches) {
		return matches.length;
	} else {
		return 0;
	}
}

// Replace 'mm-ddStamp' with timestamp using mm-dd and return # matches replaced
function mmddStamp(tabId) {

	const tab = DocumentApp.getActiveDocument().getTab(tabId);
	const body = tab.asDocumentTab().getBody();
	const text = body.getText();

	const year = new Date().getFullYear();

	// Search for occurrences of MM-DDStamp
	const regexMMDD = /\b(\d{2})-(\d{2})Stamp\b/g;

	// Manually check for all instances and replace
	const matches = text.match(regexMMDD); // Get all matching instances

	if (matches) {
		// Loop through matches and replace them
		matches.forEach(match => {
			// Extract month and day
			const parts = match.match(/(\d{2})-(\d{2})Stamp/);
			if (parts) {
				const month = parts[1];
				const day = parts[2];
				const newDate = month + '-' + day + '-' + year + '-00-00';

				// Replace the match in the document
				body.replaceText(match, newDate);
			}
		});
	}

	if (matches) {
		return matches.length;
	} else {
		return 0;
	}
}