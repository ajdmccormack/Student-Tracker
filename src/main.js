'use strict';

class Drive {
    static copy(fileId, resource) {
        return gapi.client.drive.files.copy({
            fileId: fileId,
            resource: resource
        }).then();
    }

    static create(resource) {
        return gapi.client.drive.files.create({
            resource: resource
        }).then();
    }

    static createFolder(name) {
        return gapi.client.drive.files.create({
            resource: {
                name: name,
                mimeType: 'application/vnd.google-apps.folder'
            }
        }).then();
    }

    static delete(fileId) {
        return gapi.client.drive.files.delete({
            fileId: fileId
        }).then();
    }

    static get(fileId, fields = 'id, kind, mimeType, name') {
        return gapi.client.drive.files.get({
            fileId: fileId,
            fields: fields
        }).then(function (obj) {
            return obj.result;
        });
    }

    static list(q = '') {
        return gapi.client.drive.files.list({
            q: q
        }).then(function (obj) {
            return obj.result.files;
        });
    }
}

class Spreadsheet {
    constructor (spreadsheetId) {
        this._spreadsheetId = spreadsheetId;
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
            spreadsheetId: this._spreadsheetId,
            range: sheet + '!A:A',
            valueInputOption: 'USER_ENTERED',
            values: body
        }).then();
    }

    getValues(sheet, range) {
        return gapi.client.sheets.spreadsheets.values.get({
            spreadsheetId: this._spreadsheetId,
            range: sheet + '!' + range
        }).then();
    }

    getSpreadsheet() {
        return gapi.client.sheets.spreadsheets.get({
            spreadsheetId: this._spreadsheetId
        }).then();
    }

    getSheets() {
        return this.getSpreadsheet().then(function (obj) {
            return obj.result.sheets;
        });
    }

    setValues(sheet, range, body) {
        return gapi.client.sheets.spreadsheets.values.update({
            spreadsheetId: this._spreadsheetId,
            range: sheet + '!' + range,
            valueInputOption: 'USER_ENTERED',
            values: body
        }).then();
    }

    sort() {
        return gapi.client.sheets.spreadsheets.batchUpdate({
            spreadsheetId: this._spreadsheetId,
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

class ComponentContainer {
    constructor () {
        this._components = [];
    }

    add(component) {
        this._components[this._components.length] = component;
    }

    get(index) {
        return this._components[index];
    }

    getByType(type) {
        var components = [];

        this._components.forEach(function (e) {
            if (e.types.includes(type)) {
                components[components.length] = e;
            }
        });

        return components;
    }

    update() {
        var lock = false;

        for (var component of this._components) {
            if (!lock) {
                component.enable();
            } else {
                component.disable();
            }

            if (component instanceof Assessment && !component.isCompleted) {
                lock = true;
            }
        }
    }

    disableAll() {
        this._components.forEach(function (component) {
            component.disable();
        });
    }

    get size() {
        return this._components.length;
    }
}

class Component {
	constructor (index, element, id) {
		this._index = index;
	    this._element = element;

        this._types = ['Component'];

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

    get types() {
        return this._types;
    }
}

class Task extends Component {
    constructor (index, element, id) {
    	super(index, element, id);

        this._types[this._types.length] = 'Task';

        this._isCompleted = false;
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

        this._types[this._types.length] = 'Assignment';

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

        this._types[this._types.length] = 'Assessment';

        this._id = -1;
    }

    enable() {
        console.log('Enabling Assessment ' + this._index);

        super.enable();

        if (!this._isCompleted) {
            this.onUpdate();
            
            this._id = setInterval(this.onUpdate.bind(this), 30000);

            console.log('Assessment ' + this._index + ' interval id: ' + this._id);
        }
    }

    disable() {
        console.log('Disabling Assessment ' + this._index);

        super.disable();

        clearInterval(this._id);
    }

    onUpdate() {
        console.log('Updating Assessment ' + this._index);

        FORM_SPREADSHEET.getSheets().then(function (sheets) {
            var sheet = sheets[this._assessmentIndex];

            console.log('Assessment sheet: ' + sheet.properties.title);
            
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

                    console.log('Assessment ' + this._index + ': found student and is completed');

                    break;
                }
            }

            if (success && !this._isCompleted || !success && this._isCompleted) {
                student.getIndex().then(function (studentIndex) {
                    MAIN_SPREADSHEET.setValues(SHEET, Spreadsheet.getLetter(this._index + 2) + (studentIndex + 2), [[success ? '5' : '']]);
                }.bind(this));

                this._isCompleted = success;

                clearInterval(this._id);

                COMPONENT_CONTAINER.update();
            }

            if (!success) {
                console.log('Assessment ' + this._index + ': did not find student or is not completed');
            }
        }.bind(this));
    }
}

class Clone extends Component {
    constructor (index, element, id) {
        super(index, element, CLONE_ID);

        this._fileId = this._element.href.match(/(?:\/)(.{44})(?:\/)/)[1];
    }

    enable() {
        super.enable();

        this._element.onclick = function () {

        };
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

const DISCOVERY_DOCS = ['https://sheets.googleapis.com/$discovery/rest?version=v4', 'https://www.googleapis.com/discovery/v1/apis/drive/v3/rest']
const SCOPES = 'https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive';

const SCRIPT = document.createElement('script');

const MAIN_SPREADSHEET = new Spreadsheet(MAIN_SPREADSHEET_ID);
const FORM_SPREADSHEET = new Spreadsheet(FORM_SPREADSHEET_ID);

const COMPONENT_CONTAINER = new ComponentContainer();

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

	allElements.forEach(function (e) {
        if (e.tagName != 'SCRIPT') {
    		var text = '';
    		var childNode = e.firstChild;

    		while (childNode) {
    			if (childNode.nodeType == 3) {
    				text += childNode.data;
    			}

    			childNode = childNode.nextSibling;
    		}

    		if (text.includes(ASSIGNMENT_ID)) {
    			COMPONENT_CONTAINER.add(new Assignment(index, e));

    	        index++;
    		}

            if (text.includes(ASSESSMENT_ID)) {
                COMPONENT_CONTAINER.add(new Assessment(index, e, assessmentIndex));

                index++;
                assessmentIndex++;
            }

            if (text.includes(CLONE_ID)) {
                COMPONENT_CONTAINER.add(new Clone(index, e));

                index++;
            }
        }
	});

	COMPONENT_CONTAINER.disableAll();

	document.body.appendChild(SCRIPT);
}

function handleClientLoad() {
    gapi.load('client:auth2', initClient);
}

function initClient() {
    gapi.client.init({
        apiKey: API_KEY,
        discoveryDocs: DISCOVERY_DOCS,
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
            Drive.get(MAIN_SPREADSHEET_ID, 'ownedByMe').then(function (file) {
                if (!file.ownedByMe) {
                    Drive.delete(MAIN_SPREADSHEET_ID);
                }
            });

        	if (data != null) {
                var tasks = COMPONENT_CONTAINER.getByType('Task');
        		
                data.forEach(function (value, i) {
        			if (i < tasks.length && value == '5') {
        				tasks[i].setCompleted(true);
        			}
        		});
        	} else {
        		var period = -1;

        		do {
                    period = prompt(PROMPT_PERIOD);

                    if (period === null) {
                        gapi.auth2.getAuthInstance().signOut();

                        return;
                    }
                } while (!period == 'admin' && (!Number.isInteger(period = Number(period)) || (period < 1 || period > 7)));

        		student.period = period;

        		MAIN_SPREADSHEET.appendValues(SHEET, [[student.period, student.name]]).then(function () {
                    MAIN_SPREADSHEET.sort();
                });
        	}

        	COMPONENT_CONTAINER.update();
        });

        buttonSignIn.style.display = 'none';
        buttonSignOut.style.display = '';
    } else {
        COMPONENT_CONTAINER.disableAll();

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
