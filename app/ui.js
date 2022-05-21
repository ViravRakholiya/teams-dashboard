// Select DOM elements to work with
const welcomeDiv = document.getElementById("WelcomeMessage");
const signInButton = document.getElementById("SignIn");
const cardDiv = document.getElementById("card-div");
const mailButton = document.getElementById("readMail");
const profileButton = document.getElementById("seeProfile");
const profileDiv = document.getElementById("profile-div");

function showWelcomeMessage(username) {
    // Reconfiguring DOM elements
    cardDiv.style.display = 'initial';
    welcomeDiv.innerHTML = `Welcome ${username}`;
    signInButton.setAttribute("onclick", "signOut();");
    signInButton.setAttribute('class', "btn btn-success")
    signInButton.innerHTML = "Sign Out";
}

function updateUI(data, endpoint,) {
    console.log('Graph API responded at: ' + new Date().toString());

    if (endpoint === graphConfig.graphMeEndpoint) {
        profileDiv.innerHTML = ''
        const title = document.createElement('p');
        title.innerHTML = "<strong>Title: </strong>" + data.jobTitle;
        const email = document.createElement('p');
        email.innerHTML = "<strong>Mail: </strong>" + data.mail;
        const phone = document.createElement('p');
        phone.innerHTML = "<strong>Phone: </strong>" + data.businessPhones[0];
        const address = document.createElement('p');
        address.innerHTML = "<strong>Location: </strong>" + data.officeLocation;
        profileDiv.appendChild(title);
        profileDiv.appendChild(email);
        profileDiv.appendChild(phone);
        profileDiv.appendChild(address);

    } else if (endpoint === graphConfig.graphMailEndpoint) {
        if (!data.value) {
            alert("You do not have a mailbox!")
        } else if (data.value.length < 1) {
            alert("Your mailbox is empty!")
        } else {

            data.value.map((d, i) => {
                // Keeping it simple
                if (i < 10) {
                    const listItem = document.createElement("a");
                    listItem.setAttribute("class", "list-group-item list-group-item-action")
                    listItem.setAttribute("id", "list" + i + "list")
                    listItem.setAttribute("data-toggle", "list")
                    listItem.setAttribute("href", "#list" + i)
                    listItem.setAttribute("role", "tab")
                    listItem.setAttribute("aria-controls", i)
                    listItem.innerHTML = d.subject;
                    tabList.appendChild(listItem)

                    const contentItem = document.createElement("div");
                    contentItem.setAttribute("class", "tab-pane fade")
                    contentItem.setAttribute("id", "list" + i)
                    contentItem.setAttribute("role", "tabpanel")
                    contentItem.setAttribute("aria-labelledby", "list" + i + "list")
                    contentItem.innerHTML = "<strong> from: " + d.from.emailAddress.address + "</strong><br><br>" + d.bodyPreview + "...";
                    tabContent.appendChild(contentItem);
                }
            });
        }
    } else if(endpoint == graphConfig.graphUsersEndpoint){
        if (!data.value) {
            alert("You do not have a any User!")
        } else if (data.value.length < 1) {
            alert("You do not have a any User!")
        }else{
            const tabContent = document.getElementById("nav-tabContent");
            const tabList = document.getElementById("list-tab");
            tabList.innerHTML = ''; // clear tabList at each readMail call
            let table = `
            <div class="row mt-2">
                <div class="col-md-3">
                    <select class="form-control mt-2" id="dpdfilteruser">
                        <option value="0" selected>All Users</option>
                        <option value="1">By Team Licence</option>
                    </select>
                </div>
                <div class="col-md-3">
                    <input class="form-control mt-2" id="searchUser" placeholder="Search" />
                </div>
            </div>
            <div class="table-responsive mt-2" style="height:70vh;">
                <table class="table">
                    <thead>
                        <tr>
                            <th>#</th>
                            <th>Name</th>
                            <th>Mail</th>
                            <th>Principal Name</th>
                        </tr>
                    </thead>
                <tbody id="t-data">`
            count=1;
            data.value.map((d, i) => {
                table = table +
                `<tr>
                    <th>${count++}</th>
                    <td>${d.displayName}</td>
                    <td>${d.mail}</td>
                    <td>${d.userPrincipalName}</td>
                </tr>`
            });

            table = table +
            `</tbody>
            </table>
        </div> `

        tabList.innerHTML = table;


        const txtsearchUser = document.getElementById("searchUser")
        txtsearchUser.onchange = function(){
            searchUser(this.value)
        }
        const dpdfilteruser = document.getElementById("dpdfilteruser")
        dpdfilteruser.onchange = function(){
            filterUsers(this,data)
        }
        }
    }else if(endpoint == graphConfig.graphUsersStatus){
        if (!data) {
            alert("You do not have a any User!")
        } else if (data.length < 1) {
            alert("You do not have a any User!")
        }else{
            const tabContent = document.getElementById("nav-tabContent");
            const tabList = document.getElementById("list-tab");
            tabList.innerHTML = ''; // clear tabList at each readMail call
            //console.log(data)
            let table = `
            <div class="row mt-2">
                <div class="col-md-3">
                    <input class="form-control mt-2" id="searchUser" placeholder="Search" />
                </div>
                <div class="col-md-1">
                    <button class="form-control mt-2" id="refreshStatus">
                        <i class="fa fa-refresh"></i>
                    </button>
                </div>
            </div>
            <div class="table-responsive mt-2" style="height:70vh;">
                <table class="table">
                    <thead>
                        <tr>
                            <th>#</th>
                            <th>Name</th>
                            <th>Mail</th>
                            <th>Availability</th>
                            <th>Schedule</th>
                        </tr>
                    </thead>
                <tbody id="t-data">`
            count=1;
            data.map((d, i) => {
                table = table +
                `<tr>
                    <th>${count++}</th>
                    <td>${d.displayName}</td>
                    <td>${d.mail}</td>
                    <td>${d.availability}</td>
                    <td><buttton onclick="seeUserSchedule('${d.id}');"><i class="fa fa-calendar"></i></button></td>
                </tr>`
            });''

            table = table +
            `</tbody>
            </table>
            </div> `

            tabList.innerHTML = table;
            const txtsearchUser = document.getElementById("searchUser")
            txtsearchUser.onchange = function(){
                viewUsersStatus('search',this.value)
            }

            const refreshBtn = document.getElementById("refreshStatus")
            refreshBtn.onclick = function(){
                viewUsersStatus()
            }
            
        }
    } else if(endpoint == graphConfig.graphGetGroupEndPoint){
        if (!data.value) {
            alert("You do not have a any User!")
        } else if (data.value.length < 1) {
            alert("You do not have a any User!")
        }else{
            const tabContent = document.getElementById("nav-tabContent");
            const tabList = document.getElementById("list-tab");
            tabList.innerHTML = ''; // clear tabList at each readMail call
            let table = `
            <div class="table-responsive mt-2" style="height:70vh;">
                <table class="table">
                    <thead>
                        <tr>
                            <th>#</th>
                            <th>Group Name</th>
                            <th>View</th>
                        </tr>
                    </thead>
                <tbody id="t-data">`
            count=1;
            data.value.map((d, i) => {
                table = table +
                `<tr>
                    <th>${count++}</th>
                    <td>${d.displayName}</td>
                    <td><buttton onclick="seeparticularGroup('${d.id}');"><i class="fa fa-eye"></i></button></td>
                </tr>`
            });

            table = table +
            `</tbody>
            </table>
        </div> `

        tabList.innerHTML = table;
        }
    }
}

function headerUiFOrFilter(count,data){
    const table = document.getElementById("t-data");
    if(count == 1){
        table.innerHTML = '';
    }
    var row = table.insertRow(-1);
    var cellcount = row.insertCell(0);
    var celldisplayName = row.insertCell(1);
    var cellmail = row.insertCell(2);
    var userPrincipalName = row.insertCell(3);
    var cellid = row.insertCell(4);
    cellcount.innerHTML = count
    celldisplayName.innerHTML = data.displayName
    cellmail.innerHTML = data.mail
    userPrincipalName.innerHTML = data.userPrincipalName
    cellid.innerHTML = data.id
}

function displaySchedule(scheduleDetails){
    const displayData =  document.getElementById("mymodalcontent");
    displayData.innerHTML = ''
    if (!scheduleDetails || scheduleDetails.length < 1) {
        displayData.innerHTML = '<h1>No appointments schedule for next 24 hours</h1>'
    }else{
        let table = `<div class="table-responsive mt-2" style="height:70vh;">
        <table class="table">
            <thead>
                <tr>
                    <th>Sr. No.</th>
                    <th>Subject</th>
                    <th>Location</th>
                    <th>Start Date</th>
                    <th>Start Time</th>
                    <th>End Date</th>
                    <th>End Time</th>
                </tr>
            </thead>
        <tbody>`
        count=1;
        scheduleDetails.map((d, i) => {
            table = table +
            `<tr>
                <th>${count++}</th>
                <td>${d.subject}</td>
                <td>${d.location["displayName"]}</td>
                <td>${getDate(new Date(d.start["dateTime"]))}</td>
                <td>${getTime(new Date(d.start["dateTime"]))}</td>
                <td>${getDate(new Date(d.end["dateTime"]))}</td>
                <td>${getTime(new Date(d.end["dateTime"]))}</td>
            </tr>`
        });''

        table = table +
        `</tbody>
        </table>
        </div> `

        displayData.innerHTML = table;
    }
    $('#mymodal').modal('show');
}