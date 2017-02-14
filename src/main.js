'use strict';

class Spreadsheet {
    constructor (spreadsheetId) {
        this.spreadsheetId = spreadsheetId;
    }

    static getLetter(index) {
        var letters = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];
        var letter = index == 0 ? letters[0] : '';

        for (var i = index; i > 0; i = Math.floor(i / 26)) {
            var digit = Math.floor(i % 26);

            letter = (digit == 1 && i != index ? letters[0] : letters[digit]) + letter;
        }

        return letter;
    };

    appendValues(sheet, body) {
        return gapi.client.sheets.spreadsheets.values.append({
            spreadsheetId: this.spreadsheetId,
            range: sheet + '!A:A',
            valueInputOption: 'USER_ENTERED',
            values: body
        }).then();
    }

    getValues(sheet, range) {
        return gapi.client.sheets.spreadsheets.values.get({
            spreadsheetId: this.spreadsheetId,
            range: sheet + '!' + range
        }).then();
    }

    getSpreadsheet() {
        return gapi.client.sheets.spreadsheets.get({
            spreadsheetId: this.spreadsheetId
        }).then();
    }

    getSheets() {
        return this.getSpreadsheet().then(function (obj) {
            return obj.result.sheets;
        });
    }

    setValues(sheet, range, body) {
        return gapi.client.sheets.spreadsheets.values.update({
            spreadsheetId: this.spreadsheetId,
            range: sheet + '!' + range,
            valueInputOption: 'USER_ENTERED',
            values: body
        }).then();
    }

    sort() {
        return gapi.client.sheets.spreadsheets.batchUpdate({
            spreadsheetId: this.spreadsheetId,
            requests: [{
                sortRange: {
                    range: {
                        sheetId: 0,
                        startRowIndex: 1,
                        endRowIndex: 1000,
                        startColumnIndex: 0,
                        endColumnIndex: 1000
                    },
                    sortSpecs: [{
                        dimensionIndex: 0,
                        sortOrder: 'ASCENDING'
                    }, {
                        dimensionIndex: 1,
                        sortOrder: 'ASCENDING'
                    }]
                }
            }]
        }).then();
    }
}

class TaskContainer {
    constructor () {
        this._tasks = [];
    }

    add(task) {
        this._tasks[this._tasks.length] = task;
    }

    get(index) {
        return this._tasks[index];
    }

    update() {
        var lock = false;

        for (var task of this._tasks) {
            if (!lock) {
                task.enable();
            } else {
                task.disable();
            }

            if (task instanceof Assessment && !task.isCompleted) {
                lock = true;
            }
        }
    }

    disableAll() {
        this._tasks.forEach(function (task) {
            task.disable();
        });
    }

    get size() {
        return this._tasks.length;
    }
}

class Task {
    constructor (index, element, id) {
        this._index = index;
        this._element = element;

        this._isCompleted = false;

        this._allChildren = function findAllChildren(parent, children) {
            if (parent.childElementCount == 0) {
                return children;
            } else {
                Array.from(parent.children).forEach(function (child) {
                    children[children.length] = child;

                    findAllChildren(child, children);
                });
            }

            return children;
        }(this._element, []);

        for (var node of this._element.childNodes) {
            if (node.nodeType == 3 && node.data.includes(id)) {
                node.data = node.data.replace(id, '');

                break;
            }
        }
    }

    enable() {
        this._element.onclick = function () {};
        this._element.style.setProperty('color', '');

        this._allChildren.forEach(function (element) {
            element.onclick = function () {};
            element.style.setProperty('color', '');
        });
    }

    disable()  {
        this._element.onclick = function (e) {
            e.preventDefault();
        };
        this._element.style.setProperty('color', 'gray', 'important');

        this._allChildren.forEach(function (element) {
            element.onclick = function (e) {
                e.preventDefault();
            };
            element.style.setProperty('color', 'gray', 'important');
        });
    }

    get isCompleted() {
        return this._isCompleted;
    }

    setCompleted(isCompleted) {
        this._isCompleted = isCompleted;
    }
}

class Assignment extends Task {
    constructor (index, element) {
        super(index, element, ASSIGNMENT_ID);

        this._checkbox = document.createElement('input');
        this._checkbox.type = 'checkbox';
        this._checkbox.disabled = true;
        this._checkbox.onchange = function () {
            var value = this._checkbox.checked ? '5' : '';

            student.getIndex().then(function (studentIndex) {
                MAIN_SPREADSHEET.setValues(SHEET, Spreadsheet.getLetter(this._index + 2) + (studentIndex + 2), [[value]]);
            }.bind(this));

            this._isCompleted = this._checkbox.checked;
        }.bind(this);

        this._element.insertBefore(document.createTextNode(' '), this._element.firstChild);
        this._element.insertBefore(this._checkbox, this._element.firstChild);
    }

    enable() {
        super.enable();

        this._checkbox.disabled = false;
    }

    disable() {
        super.disable();

        this._checkbox.disabled = true;
    }

    setCompleted(isCompleted) {
        super.setCompleted(isCompleted);

        this._checkbox.checked = isCompleted;
    }
}

class Assessment extends Task {
    constructor (index, element, assessmentIndex) {
        super(index, element, ASSESSMENT_ID);

        this._assessmentIndex = assessmentIndex;

        this._id = -1;
    }

    enable() {
        super.enable();

        clearInterval(this._id);
        this.onUpdate();

        this._id = setInterval(this.onUpdate.bind(this), 5000);
    }

    disable() {
        super.disable();

        clearInterval(this._id);
    }

    onUpdate() {
        FORM_SPREADSHEET.getSheets().then(function (sheets) {
            var sheet = sheets[(sheets.length - 1) - this._assessmentIndex];
            
            return FORM_SPREADSHEET.getValues(sheet.properties.title, 'B1:' + Spreadsheet.getLetter(sheet.properties.gridProperties.columnCount - 1));
        }.bind(this)).then(function (obj) {
            var values = obj.result.values;
            var success = false;

            var emailIndex = -1;
            var scoreIndex = -1;

            for (var i = 0; i < values[0].length; i++) {
                if (emailIndex != -1 && scoreIndex != -1) {
                    break;
                }

                switch (values[0][i]) {
                    case 'Email Address':
                        emailIndex = i;

                        break;

                    case 'Score':
                        scoreIndex = i;

                        break;
                }
            }

            values.splice(0, 1);
            
            for (var row of values) {
                if (row[emailIndex] == student.email && Number(row[scoreIndex].substring(0, row[scoreIndex].indexOf('/') - 1)) == Number(row[scoreIndex].substring(row[scoreIndex].indexOf('/') + 2))) {
                    success = true;

                    break;
                }
            }

            if (success && !this._isCompleted || !success && this._isCompleted) {
                student.getIndex().then(function (studentIndex) {
                    MAIN_SPREADSHEET.setValues(SHEET, Spreadsheet.getLetter(this._index + 2) + (studentIndex + 2), [[success ? '5' : '']]);
                }.bind(this));

                this._isCompleted = success;

                TASK_CONTAINER.update();
            }
        }.bind(this));
    }
}

class Student {
    constructor (rawName, email) {
        this._rawName = rawName;
        this._email = email;

        this._name = this._rawName.substring(this._rawName.lastIndexOf(' ') + 1) + ', ' + this._rawName.substring(0, this._rawName.lastIndexOf(' '));
    }

    getIndex() {
        return MAIN_SPREADSHEET.getValues(SHEET, 'B2:B').then(function (obj) {
            var names = obj.result.values != undefined ? obj.result.values : [[]];

            for (var i = 0; i < names.length; ++i) {
                if (names[i][0] == student.name) {
                    return i;
                }
            }

            return -1;
        });
    }

    getData() {
        return this.getIndex().then(function (index) {
            return index != -1 ? MAIN_SPREADSHEET.getValues(SHEET, 'C' + (index + 2) + ':' + (index + 2)) : null;
        }).then(function (obj) {
            return obj != null && obj.result.values != undefined ? obj.result.values[0] : obj != null ? [[]] : null;
        });
    }

    get name() {
        return this._name;
    }

    get email() {
        return this._email;
    }

    get period() {
        return this._period;
    }

    set period(period) {
        this._period = period;
    }
}

const DISCOVERY_URL = 'https://sheets.googleapis.com/$discovery/rest?version=v4'
const SCOPES = 'https://www.googleapis.com/auth/spreadsheets';

const SCRIPT = document.createElement('script');

const MAIN_SPREADSHEET = new Spreadsheet(MAIN_SPREADSHEET_ID);
const FORM_SPREADSHEET = new Spreadsheet(FORM_SPREADSHEET_ID);

const TASK_CONTAINER = new TaskContainer();

const ELEMENTS = [];
const CHECKBOXES = [];

var student;

var buttonSignIn;
var buttonSignOut;

main();

function main() {
    init();
    run();
}

function init() {
    SCRIPT.src = 'https://apis.google.com/js/api.js';
    SCRIPT.onload = function() {
        handleClientLoad();
    };
    SCRIPT.defer = 'defer';
    SCRIPT.async = 'async';
}

function run() {
	var allElements = Array.from(document.getElementsByTagName('*'));
    var index = 0;
	var assessmentIndex = 0;

	allElements.forEach(function (element) {
		var text = '';
		var childNode = element.firstChild;

		while (childNode) {
			if (childNode.nodeType == 3) {
				text += childNode.data;
			}

			childNode = childNode.nextSibling;
		}

		if (text.includes(ASSIGNMENT_ID)) {
			TASK_CONTAINER.add(new Assignment(index, element));

	        index++;
		} else if (text.includes(ASSESSMENT_ID)) {
            TASK_CONTAINER.add(new Assessment(index, element, assessmentIndex));

            index++;
            assessmentIndex++;
        }
	});

	TASK_CONTAINER.disableAll();

	document.body.appendChild(SCRIPT);
}

function handleClientLoad() {
    gapi.load('client:auth2', initClient);
}

function initClient() {
    gapi.client.init({
        apiKey: API_KEY,
        discoveryDocs: [DISCOVERY_URL],
        clientId: CLIENT_ID,
        scope: SCOPES
    }).then(function () {
        buttonSignIn = document.createElement('li');
        buttonSignIn.className = 'wsite-menu-item-wrap   wsite-nav-6';
        buttonSignIn.style = 'position: relative;';
        buttonSignIn.innerHTML = '<a id="authorize-button" class="wsite-button wsite-button-small wsite-button-normal" onclick="handleSignInClick()"><span class="wsite-button-inner">Sign In</span></a>';

        buttonSignOut = buttonSignIn.cloneNode(true);
        buttonSignOut.className = 'wsite-menu-item-wrap   wsite-nav-7';
        buttonSignOut.innerHTML = '<a id="signout-button" class="wsite-button wsite-button-small wsite-button-normal" onclick="handleSignOutClick()"><span class="wsite-button-inner">Sign Out</span></a>';

        var appendButtons = function () {
            if (document.documentElement.clientWidth <= 1024) {
                document.getElementsByClassName('wsite-menu-default wsite-menu-slide')[0].appendChild(buttonSignIn);
                document.getElementsByClassName('wsite-menu-default wsite-menu-slide')[0].appendChild(buttonSignOut);
            } else {
                document.getElementsByClassName('wsite-menu-default')[0].appendChild(buttonSignIn);
                document.getElementsByClassName('wsite-menu-default')[0].appendChild(buttonSignOut);
            }
        };

        appendButtons();

        window.onresize = appendButtons;

        updateSignInStatus(gapi.auth2.getAuthInstance().isSignedIn.get());
        gapi.auth2.getAuthInstance().isSignedIn.listen(updateSignInStatus);
    });
}

function updateSignInStatus(isSignedIn) {
    if (isSignedIn) {
        var rawName = gapi.auth2.getAuthInstance().currentUser.get().getBasicProfile().getName();
        var email = gapi.auth2.getAuthInstance().currentUser.get().getBasicProfile().getEmail();

        student = new Student(rawName, email);

        student.getData().then(function (data) {
        	if (data != null) {
        		data.forEach(function (value, i) {
        			if (i < TASK_CONTAINER.size && value == '5') {
        				TASK_CONTAINER.get(i).setCompleted(true);
        			}
        		});
        	} else {
        		var period = -1;

        		do {
        			period = prompt('In which class period are you? (1, 2, 4, 5, 7)', '');

        			if (period === null) {
        				gapi.auth2.getAuthInstance().signOut();

        				return;
        			} else {
        				period = Number(period);
        			}
        		} while (!Number.isInteger(period) || (period < 1 || period > 7));

        		student.period = period;

        		MAIN_SPREADSHEET.appendValues(SHEET, [[student.period, student.name]]).then(function () {
                    MAIN_SPREADSHEET.sort();
                });
        	}

        	TASK_CONTAINER.update();
        });

        buttonSignIn.style.display = 'none';
        buttonSignOut.style.display = '';
    } else {
        TASK_CONTAINER.disableAll();

        buttonSignOut.style.display = 'none';
        buttonSignIn.style.display = '';
    }
}

function handleSignInClick(event) {
    gapi.auth2.getAuthInstance().signIn();
}

function handleSignOutClick(event) {
    gapi.auth2.getAuthInstance().signOut();
}
