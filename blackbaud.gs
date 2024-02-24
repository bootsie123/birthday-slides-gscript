const BLACKBAUD_CLIENT_ID = "REPLACE_WITH_BLACKBAUD_CLIENT_ID";
const BLACKBAUD_CLIENT_SECRET = "REPLACE_WITH_BLACKBAUD_CLIENT_SECRET";

const BLACKBAUD_SUBSCRIPTION = "REPLACE_WITH_BLACKBAUD_SUBSCRIPTION"

const API_URL = "https://api.sky.blackbaud.com";

/**
 * Creates a new instance of the BlackbaudService
 * 
 * @return {OAuth2Service} An OAuth2Service object
 */
function getBlackbaudService_() {
  return OAuth2.createService("blackbaud")
    .setAuthorizationBaseUrl("https://oauth2.sky.blackbaud.com/authorization")
    .setTokenUrl("https://oauth2.sky.blackbaud.com/token")
    .setClientId(BLACKBAUD_CLIENT_ID)
    .setClientSecret(BLACKBAUD_CLIENT_SECRET)
    .setCallbackFunction("authCallback_")
    .setPropertyStore(PropertiesService.getUserProperties())
    .setLock(LockService.getUserLock())
    .setTokenHeaders({
      "Authorization": "Basic " + Utilities.base64Encode(`${BLACKBAUD_CLIENT_ID}:${BLACKBAUD_CLIENT_SECRET}`)
    });
}

/**
 * Handles first time Blackbaud API authorization through an HTML template served via
 * a Google App Scripts web deployment
 */
function doGet() {
  const blackbaud = getBlackbaudService_();

  const template = HtmlService.createTemplateFromFile("setup");
  
  template.setup = blackbaud.hasAccess();
  template.redirectUri = OAuth2.getRedirectUri();
  template.authorizationUrl = blackbaud.getAuthorizationUrl();
  
  return template.evaluate();
}

/**
 * Handles an OAuth2 callback and serves a page with the accompanying status
 */
function authCallback_(request) {
  const template = HtmlService.createTemplateFromFile("callback");

  try {
    const blackbaud = getBlackbaudService_();

    const isAuthorized = blackbaud.handleCallback(request);

    if (isAuthorized) {
      template.title = "Success!";
      template.status = "success";
      template.message = "Account linked successfully!";
    } else {
      template.title = "Unauthorized!";
      template.status = "danger";
      template.message = "An unknown error has occured while trying to link your account. Please try again";
    }
  } catch (err) {
      template.title = "Error!";
      template.status = "danger";
      template.message = `An error has occured while linking your account: ${err.message}`;
  }

  template.status = "has-text-" + template.status;

  return template.evaluate();
}

/**
 * Fetches and parses the JSON returned from a request to the SKY API
 * 
 * @param {OAuth2Service} blackbaud A Blackbaud OAuth2Service object
 * @param {string}        url       The SKY API URL to fetch
 * @return {dict}         A dictionary corresponding to the parsed JSON
 */
function fetch_(blackbaud, url) {
  const res = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: "Bearer " + blackbaud.getAccessToken(),
      "Bb-Api-Subscription-Key": BLACKBAUD_SUBSCRIPTION
    }
  });

  return JSON.parse(res);
}

/**
 * Retrieves a list from Blackbaud
 * 
 * @param {OAuth2Service} blackbaud A Blackbaud OAuth2Service object
 * @param {string}        listId    The ID of the list to fetch
 * @param {number}        page      The page of the list to get (defaults to 1)
 * @param {number}        pageSize  The page size when getting the list (defaults to 1000)
 * @return {dict}         A dictionary containing the requested list
 */
function getList(blackbaud, listId, page = 1, pageSize = 1000) {
  return fetch_(blackbaud,`${API_URL}/school/v1/lists/advanced/${listId}?page=${page}&page_size=${pageSize}`);
}

/**
 * Retrieves information about a user from Blackbaud
 * 
 * @param {OAuth2Service} blackbaud A Blackbaud OAuth2Service object
 * @param {string}        userId    The ID of the user to retrieve
 * @return {dict}         A dictionary containing info about the requested user
 */
function getUser(blackbaud, userId) {
  return fetch_(blackbaud,`${API_URL}/school/v1/users/${userId}`);
}
