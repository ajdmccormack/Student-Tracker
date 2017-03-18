'use strict';

const API_KEY = 'AIzaSyD0kTBtlNzVae3u1LYcjZKrre563mcxsVo';
const CLIENT_ID = '627343722694-fni1on670josisrndul45j23n0bim9a5.apps.googleusercontent.com';
const MAIN_SPREADSHEET_ID = '1IIjBsBJrkgaGORKvsXruJpk-6xKC3clPnW_-AyqpUMw';
const FORM_SPREADSHEET_ID = '1bJRdFCGP2yocIVXw6ewsDBALeUkYfrVoorf5DnBpAc4';
const SHEET = 'Sheet1';
const PROMPT_PERIOD = 'In which class period are you? (1, 2, 4, 5, 7)';
const ASSIGNMENT_ID = '~';
const ASSESSMENT_ID = '`';
const CLONE_ID = '^';

class Drive {
    constructor (childFolderId) {
        this._childFolderId = childFolderId;
    }

    get childFolderId() {
        return this._childFolderId;
    }
}

Drive.Files = class {
    static copy(fileId, resource) {
        return gapi.client.drive.files.copy({
            fileId: fileId,
            resource: resource
        }).then(function (obj) {
            return obj.result;
        });
    }

    static create(resource) {
        return gapi.client.drive.files.create({
            resource: resource
        }).then(function (obj) {
            return obj.result;
        });
    }

    static createFolder(name, resource) {
        return Drive.Files.create(Object.assign({
            name: name,
            mimeType: 'application/vnd.google-apps.folder'
        }, resource));
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

Drive.Permissions = class {
    static create(fileId, resource) {
        gapi.client.drive.permissions.create({
            fileId: fileId,
            resource: resource
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
        var chain = new Promise(function (resolve, reject) {
            resolve();
        });

        for (let i = 0, count = 0; i < this._components.length; i++) {
            const component = this._components[i];

            if (!component.isEnabled) {
                chain = chain.then(function () {
                    var ret = component.enable();

                    if (count % 3 == 0) {
                        return ret;
                    }
                });

                if (component instanceof Clone) {
                    count++;
                }
            }

            if (component instanceof Assessment && !component.isCompleted) {
                return;
            }
        }
    }

    onSignIn() {
        var clones = this.getByType('Clone');

        if (clones.length != 0) {
            const HOSTNAME = window.location.hostname.match(/\.(.+)\./)[1];
            const PATHNAME = window.location.pathname.match(/\/([^.]*)/)[1];

            var parent = Drive.Files.list("mimeType = 'application/vnd.google-apps.folder' and properties has {key = 'id' and value = '" + HOSTNAME + "'}").then(function (files) {
                if (files.length != 0) {
                    console.log('CLONE: Found parent');

                    return files[0].id;
                } else {
                    console.log('CLONE: Not Found parent');

                    var folder = Drive.Files.createFolder(HOSTNAME, {
                        properties: {
                            id: HOSTNAME
                        }
                    });

                    var emailAddress = Drive.Files.get(MAIN_SPREADSHEET_ID, 'owners').then(function (obj) {
                        return obj.owners[0].emailAddress;
                    });

                    return Promise.all([folder, emailAddress]).then(function (values) {
                        var folder = values[0];
                        var emailAddress = values[1];

                        Drive.Permissions.create(folder.id, {
                            role: 'writer',
                            type: 'user',
                            emailAddress: emailAddress
                        });

                        return folder.id;
                    });
                }
            }.bind(this));

            var child = parent.then(function (parentFolderId) {
                return Drive.Files.list("mimeType = 'application/vnd.google-apps.folder' and '" + parentFolderId + "' in parents and properties has {key = 'id' and value = '" + PATHNAME + "'}")
            });

            drive = new Drive(Promise.all([parent, child]).then(function (values) {
                var parentFolderId = values[0];
                var files = values[1];

                if (files.length != 0) {
                    console.log('CLONE: Found child');

                    return files[0].id;
                } else {
                    console.log('CLONE: Not Found child');

                    return Drive.Files.createFolder(document.title, {
                        properties: {
                            id: PATHNAME
                        },
                        parents: [
                            parentFolderId
                        ]
                    }).then(function (file) {
                        return file.id;
                    });
                }
            }.bind(this)));
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
	constructor (node, id) {
	    this._element = node.parentElement;

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

        this._isEnabled = false;

        node.data = node.data.replace(id, '');
	}

	enable() {
        this._element.onclick = function () {};
        this._element.style.setProperty('color', '');

        this._allChildren.forEach(function (element) {
            element.onclick = function () {};
            element.style.setProperty('color', '');
        });

        this._isEnabled = true;
    }

    disable() {
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

        this._isEnabled = false;
    }

    get isEnabled() {
        return this._isEnabled;
    }

    get types() {
        return this._types;
    }
}

class Task extends Component {
    constructor (index, node, id) {
    	super(node, id);

        this._index = index;

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
    constructor (index, node) {
        super(index, node, ASSIGNMENT_ID);

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
    constructor (index, node, assessmentIndex) {
        super(index, node, ASSESSMENT_ID);

        this._assessmentIndex = assessmentIndex;

        this._types[this._types.length] = 'Assessment';

        this._id = -1;
    }

    enable() {
        console.log('ASSESSMENT ' + this._index + ': enabling');

        super.enable();

        if (!this._isCompleted) {
            this.onUpdate();
            
            this._id = setInterval(this.onUpdate.bind(this), 30000);

            console.log('ASSESSMENT ' + this._index + ': interval id ' + this._id);
        }
    }

    disable() {
        console.log('ASSESSMENT ' + this._index + ': disabling');

        super.disable();

        clearInterval(this._id);
    }

    onUpdate() {
        console.log('ASSESSMENT ' + this._index + ': updating');

        FORM_SPREADSHEET.getSheets().then(function (sheets) {
            var sheet = sheets[this._assessmentIndex];

            console.log('ASSESSMENT ' + this._index + ': sheet ' + sheet.properties.title);
            
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

                    console.log('ASSESSMENT ' + this._index + ': found student and is completed');

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
                console.log('ASSESSMENT ' + this._index + ': did not find student or is not completed');
            }
        }.bind(this));
    }
}

class Clone extends Component {
    constructor (node) {
        super(node, CLONE_ID);

        this._element = function findElement(e) {
            if (e.tagName == 'A') {
                return e;
            } else {
                for (var e of Array.from(e.children)) {
                    var ret = findElement(e);

                    if (ret != undefined) {
                        return ret;
                    }
                }
            }
        }(this._element);
        this._types[this._types.length] = 'Clone';

        this._fileId = this._element.href.match(/(?:d\/|id=)(.{44})/)[1];
    }

    enable() {
        super.enable();

        var files = drive.childFolderId.then(function (childFolderId) {
            return Drive.Files.list("'" + childFolderId + "' in parents and properties has {key = 'id' and value = '" + this._fileId + "'}");
        }.bind(this));

        var ret = Promise.all([drive.childFolderId, files]).then(function (values) {
            var childFolderId = values[0];
            var files = values[1];

            if (files.length != 0) {
                console.log('CLONE ' + this._fileId + ': Found file');

                return Drive.Files.get(files[0].id, 'webViewLink');
            } else {
                console.log('CLONE ' + this._fileId + ': Not Found file');

                return Drive.Files.copy(this._fileId, {
                    properties: {
                        id: this._fileId
                    },
                    parents: [
                        childFolderId
                    ]
                }).then(function (file) {
                    return Drive.Files.get(file.id, 'webViewLink');
                });
            }
        }.bind(this)).then(function (file) {
            this._element.href = file.webViewLink;
        }.bind(this));

        return ret;
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
var drive;

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
            for (var node = e.firstChild; node; node = node.nextSibling) {
                if (node.nodeType == 3) {
                    if (node.data.includes(ASSIGNMENT_ID)) {
                        COMPONENT_CONTAINER.add(new Assignment(index, node));

                        index++;
                    }

                    if (node.data.includes(ASSESSMENT_ID)) {
                        COMPONENT_CONTAINER.add(new Assessment(index, node, assessmentIndex));

                        index++;
                        assessmentIndex++;
                    }

                    if (node.data.includes(CLONE_ID)) {
                        COMPONENT_CONTAINER.add(new Clone(node));
                    }
                }
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

        COMPONENT_CONTAINER.onSignIn();

        student.getData().then(function (data) {
            /* Drive.Files.get(MAIN_SPREADSHEET_ID, 'ownedByMe').then(function (file) {
                if (!file.ownedByMe) {
                    Drive.Files.delete(MAIN_SPREADSHEET_ID);
                }
            }); */

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
