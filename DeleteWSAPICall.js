let groupsCounter = 0;
let deletedGroupsCounter = 0;
let emptyGroupsCounter = 0;
let emptyWorkspaceIds = [];
let accessToken;
let numOfWorkspaces = 0;
let userEmail; 

async function init(data)  {
    accessToken = data.token;
    const expandedGroupsString = data.url;
    userEmail = data.email;
    getGroups(accessToken, expandedGroupsString);
};

const getGroups = (token, groupInitString) => {
    let header = new Headers();
    header.append("Authorization", `Bearer ${token}`);
    fetch(groupInitString, {
        headers: header
    }).then(res=> {
        res.json().then(groups => {
            getReports(groups);
        })
    }).catch(() => {
        alert('Cannot fetch data, getGroups...');
    });
};

function getReports(groups){
    for (let i = 0; i < groups.value.length; i++) {
        const reports = groups.value[i].reports;
        const datasets = groups.value[i].datasets;

        groups.value[i].numOfReports = reports.length;
        groups.value[i].numOfdatasets = datasets.length;

        numOfWorkspaces++;
        AppendTable(groups.value[i]);
    }

    document.getElementById('ws-count').title = groupsCounter;
    document.getElementById('ws-empty').title = emptyGroupsCounter;
}

/**
 * Appends the workspaces to the DOM
 * @param {Array} gArray 
 */
function AppendTable(gArray) {
    groupsCounter++;
    let table = document.querySelector('.list-groups table');

    let trEle = document.createElement('tr');
    trEle.setAttribute('class', 'workspaces');
    trEle.setAttribute('id', `rownum${gArray.id}`);
    let tdEle1 = document.createElement('td');
    tdEle1.innerHTML = gArray.name;
    trEle.appendChild(tdEle1);
    let tdEle2 = document.createElement('td');
    tdEle2.innerHTML = gArray.numOfReports;
    let tdEle3 = document.createElement('td');
    tdEle3.innerHTML = gArray.state;
    let tdEle4 = document.createElement('td');
    tdEle4.innerHTML = gArray.type;
    trEle.appendChild(tdEle4);
    trEle.appendChild(tdEle3);
    trEle.appendChild(tdEle2);
    table.appendChild(trEle);

    if (gArray.numOfReports === 0 && gArray.state !== 'Deleted' && gArray.type !== 'PersonalGroup') {
        emptyGroupsCounter++
        if (gArray.numOfdatasets > 0) {
            return tdEle2.innerHTML += " but " + gArray.numOfdatasets + " datasets"
        }
        emptyWorkspaceIds.push(gArray.id);
        let button = document.createElement('button');
        button.innerText = "Delete";
        tdEle2.appendChild(button);
        button.addEventListener('click', () => { deleteWorkspace(gArray.id); });
    }


    let deleteAllButton = document.querySelector('#deleteAll');
    deleteAllButton.style.display = 'inline-block';
}

/**
 * Will give the user admin access to the workspace before deleting it.
 * @param {GUID} groupId 
 */
async function addUserAsAdmin(groupId) {
    if (!userEmail) { // If userEmail is empty then no extra permision neded for this account
        return "No email set";
    }

    await fetch(`https://api.powerbi.com/v1.0/myorg/admin/groups/${groupId}/users`, {
        method: 'POST',
        headers: {
        'Accept': 'application/json',
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${accessToken}`
        },
        body: JSON.stringify({emailAddress: userEmail, groupUserAccessRight: "Admin"})
    })
    .then(response => response.json())
    .then((data) => {
        return data
    })
    .catch(error => error);
}

/**
 * Removes the Workspace that has the incoming groupId
 * @param {GUID} groupId 
 */
function deleteWorkspace(groupId) {
    addUserAsAdmin(groupId).then(res => {
        let header = new Headers();
        header.append("Authorization", `Bearer ${accessToken}`);
        var requestOptions = {
            method: 'DELETE',
            headers: header,
            redirect: 'follow'
        };

        fetch("https://api.powerbi.com/v1.0/myorg/groups/" + groupId, requestOptions)
        .then(response => response.text())
        .then(() => {
          document.getElementById(`rownum${groupId}`).remove();
          deletedGroupsCounter++;
          document.getElementById('ws-deleted').title = deletedGroupsCounter;
        })
        .catch(error => alert("Could not delete workspace: ", error));
    })
}

function impFunctions() {
    document.getElementById('ws-deleted').title = deletedGroupsCounter;
    let deleteAllButton = document.querySelector('#deleteAll');
    deleteAllButton.style.display = 'inline-block';
    deleteAllButton.addEventListener('click', () => {
        if (emptyWorkspaceIds.length === 0) {
            return alert("No empty Workspace found");
        }

        for (let i = 0; i < emptyWorkspaceIds.length; i++) {
            deleteWorkspace(emptyWorkspaceIds[i]);
        }
    });
}

window.addEventListener("load", function () {
    // Access the form element...
    const form = document.getElementById("form-input");
    form.addEventListener("submit", function (event) {
        event.preventDefault();
        const formInput = {
            url: form.url.value,
            token: form.token.value,
            email: form.email.value
        }
        init(formInput).then(()=>{impFunctions();});
    });
});