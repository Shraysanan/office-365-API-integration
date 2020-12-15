// configuration below set from azure app created and permissions are stated======
const msalConfig = {
  auth: {
    clientId: 'your-specific client id',
    redirectUri: 'your desired redirect uri setup in azure'
  }
};

const msalRequest = {
  scopes: [
    'user.read',
    'mailboxsettings.read',
    'calendars.readwrite',
    'contacts.readwrite',
    'mail.read',
    'mail.readbasic'
  ]
}

//================================================================================


//==============sign in authentication and sign out logic below===================
const msalClient = new msal.PublicClientApplication(msalConfig);
async function signIn() {
  // Login
  try {
    // Use MSAL to login
    const authResult = await msalClient.loginPopup(msalRequest);
    console.log('id_token acquired at: ' + new Date().toString());
    // Save the account username, needed for token acquisition
    sessionStorage.setItem('msalAccount', authResult.account.username);

    // Get the user's profile from Graph
    user = await getUser();
    // Save the profile in session
    sessionStorage.setItem('graphUser', JSON.stringify(user));
  } catch (error) {
    console.log(error);
  }
}
async function getToken() {
  let account = sessionStorage.getItem('msalAccount');
  if (!account){
    throw new Error(
      'User account missing from session. Please sign out and sign in again.');
  }

  try {
    // First, attempt to get the token silently
    const silentRequest = {
      scopes: msalRequest.scopes,
      account: msalClient.getAccountByUsername(account)
    };

    const silentResult = await msalClient.acquireTokenSilent(silentRequest);
    console.log(silentResult.accessToken);
    return silentResult.accessToken;
  } catch (silentError) {
    // If silent requests fails with InteractionRequiredAuthError,
    // attempt to get the token interactively
    if (silentError instanceof msal.InteractionRequiredAuthError) {
      const interactiveResult = await msalClient.acquireTokenPopup(msalRequest);
      return interactiveResult.accessToken;
    } else {
      throw silentError;
    }
  }
}

function signOut() {
  account = null;
  sessionStorage.removeItem('graphUser');
  msalClient.logout();
}

//================================================================================

//============gets user data and token ===========================================
const authProvider = {
  getAccessToken: async () => {
    // Call getToken in auth.js
    return await getToken();
  }
};

async function getUser() {
  return await graphClient
    .api('/me')
    // Only get the fields used by the app
    .select('id,displayName,mail,userPrincipalName,mailboxSettings')
    .get();
}
//================================================================================

//===========initialize graph client =============================================
const graphClient = MicrosoftGraph.Client.initWithMiddleware({authProvider});
//================================================================================

//=============== get contacts ===================================================
async function getcontacts(){
  const user = JSON.parse(sessionStorage.getItem('graphUser'));
  try{
    let response = await graphClient
    .api('/me/contacts')
    .get();
    console.log(response);
	} catch (error) {
	   console.log(error);
	}
}
//================================================================================
//====================create contact==============================================
const contact = {
  givenName: "demo",
  surname: "contact",
  emailAddresses: [
    {
      address: "democontact@fabrikam.onmicrosoft.com",
      name: "demo contact"
    }
  ],
  businessPhones: [
    "+1 732 555 0102"
  ]
};
async function createcontact(){
  const user = JSON.parse(sessionStorage.getItem('graphUser'));
  try{
  let response = await graphClient.api('/me/contacts').post(contact);
    console.log(response);
  } catch (error) {
     console.log(error);
  }
}
//================================================================================
//====================delete contact==============================================
async function deletecontact(){
  const user = JSON.parse(sessionStorage.getItem('graphUser'));
  try{
  // let response = await graphClient.api('/me/contacts/{id}').delete();
  let response = await graphClient.api('/me/contacts/AAMkAGU1NWU2ZjZiLTA4NTAtNDkxOC04ZDJjLTA1MzA5NzkxYzI5OQBGAAAAAABbDD4ptG68RoAg6ODJhzo8BwC4qjEIwGiuQZ7PLKncJO4JAAAAAAEOAAC4qjEIwGiuQZ7PLKncJO4JAAAI0D_7AAA=').delete();
    console.log(response);
  } catch (error) {
     console.log(error);
  }
}
//================================================================================
//=======================update contact===========================================
const updateincontact = {
  homeAddress: {
    street: "123 Some street",
    city: "Seattle",
    state: "WA",
    postalCode: "98121"
  },
  birthday: "1974-07-22"
};
async function updatecontact(){
  const user = JSON.parse(sessionStorage.getItem('graphUser'));
  try{
    let response = await graphClient.api('/me/contacts/AAMkAGU1NWU2ZjZiLTA4NTAtNDkxOC04ZDJjLTA1MzA5NzkxYzI5OQBGAAAAAABbDD4ptG68RoAg6ODJhzo8BwC4qjEIwGiuQZ7PLKncJO4JAAAAAAEOAAC4qjEIwGiuQZ7PLKncJO4JAAAI0D_7AAA=').update(updateincontact);
    // let response = await graphClient.api('/me/contacts/{id}').update(updateincontact);
    console.log(response);
  } catch (error) {
     console.log(error);
  }
}
//================================================================================
//==============get calender events===============================================
async function getEvents() {
  const user = JSON.parse(sessionStorage.getItem('graphUser'));
  let ianaTimeZone = getIanaFromWindows(user.mailboxSettings.timeZone);
  console.log(`Converted: ${ianaTimeZone}`);
  let startOfWeek = moment.tz('Asia/Calcutta').startOf('week').utc();
  let endOfWeek = moment(startOfWeek).add(7, 'day');
  try {
    let response = await graphClient
      .api('/me/calendarview')
      .header("Prefer", `outlook.timezone="India Standard Time"`)
      .query({ startDateTime: startOfWeek.format(), endDateTime: endOfWeek.format() })
      .select('subject,organizer,start,end')
      .orderby('start/dateTime')
      .top(50)
      .get();
      console.log(response);
  } catch (error) {
    console.log(error);
  }
}
//================================================================================

//========== get mails ===========================================================
async function getmails(){
  const user = JSON.parse(sessionStorage.getItem('graphUser'));
  try{
    let response = await graphClient
    .api('/me/messages')
    .get();
    console.log(response);
	} catch (error) {
      console.log(error);
	}
}
//==========buttons that decide to login or retrieve user data====================
function calendarbutton(){
  let account = sessionStorage.getItem('msalAccount');
  if(!account){
    signIn();
  }
  else{
    getEvents();
  }
}
function createcontactbutton(){
  let account = sessionStorage.getItem('msalAccount');
  if(!account){
    signIn();
  }
  else{
    createcontact();
  }
}
function updatecontactbutton(){
  let account = sessionStorage.getItem('msalAccount');
  if(!account){
    signIn();
  }
  else{
    updatecontact();
  }
}
function deletecontactbutton(){
  let account = sessionStorage.getItem('msalAccount');
  if(!account){
    signIn();
  }
  else{
    deletecontact();
  }
}
function contactbutton(){
  let account = sessionStorage.getItem('msalAccount');
  if(!account){
    signIn();
  }
  else{
    getcontacts();
  }
}
function mailbutton(){
  let account = sessionStorage.getItem('msalAccount');
  if(!account){
    signIn();
  }
  else{
    getmails();
  }
}
