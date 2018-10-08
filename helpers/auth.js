const  jwt = require('jsonwebtoken');

const credentials = {
    client: {
        id: process.env.APP_ID,
        secret: process.env.APP_PASSWORD,
    },
    auth: {
        tokenHost: 'https://login.microsoftonline.com',
        authorizePath: 'common/oauth2/v2.0/authorize',
        tokenPath: 'common/oauth2/v2.0/token'
    }
};
const oauth2 = require('simple-oauth2').create(credentials);

function getAuthUrl() {
    const returnVal = oauth2.authorizationCode.authorizeURL({
        redirect_uri: process.env.REDIRECT_URI,
        scope: process.env.APP_SCOPES
    });
    console.log(`Generated auth url: ${returnVal}`);
    return returnVal;
}

async function getTokenFromCode(auth_code, res) {
    console.log('auth code: ', auth_code);
    // var token;
    // try {
    //     let result = await oauth2.authorizationCode.getToken({
    //         code: auth_code,
    //         redirect_uri: process.env.REDIRECT_URI,
    //         scope: process.env.APP_SCOPES
    //     });
    //     token = oauth2.accessToken.create(result);
    // } catch (error) {
    //     console.log('Access Token Error', error.message);
    // }
    let result = await oauth2.authorizationCode.getToken({
        code: auth_code,
        redirect_uri: process.env.REDIRECT_URI,
        scope: process.env.APP_SCOPES
    });

    const token = oauth2.accessToken.create(result);
    console.log('Token Created', token.token);

    saveValuesToCookie(token, res);
    return token.token.access_token;
}

//Function saveToCookie <-- Token, Response

function saveValuesToCookie(token, res) {

    // Parse Identity Token
    const user = jwt.decode(token.token.id_token);

    // Save the access token in a cookie
    res.cookie('graph_access_token', token.token.access_token, {maxAge: 3600000, httpOnly: true});
    // Save the user's name in a cookie
    res.cookie('graph_user_name', user.name, {maxAge: 360000, httpOnly: true});

    // Save the refresh token in a cookie
    res.cookie('graph_refresh_token', token.token.refresh_token, {maxAge: 7200000, httpOnly: true});
    // Save the token expiration time in a cookie
    res.cookie('graph_token_expires', token.token.expires_at.getTime(), {maxAge: 3600000, httpOnly: true});

}

// Signing out router -> helper function to clear cookies

function clearCookies(res) {
    // clear cookies
    res.clearCookie('graph_access_token', {maxAge: 3600000, httpOnly: true});
    res.clearCookie('graph_user_name', {maxAge: 3600000, httpOnly: true});
    res.clearCookie('graph_refresh_token', {maxAge: 7200000, httpOnly: true});
    res.clearCookie('graph_token_expires', {maxAge: 3600000, httpOnly: true});
}

async function getAccessToken(cookies, res) {
    // Do we have an access token cached ?
    let token = cookies.graph_access_token;

    if(token) {
        // We have token, But is it expired ?

        const FIVE_MINUTES = 300000;
        const expiration = new Date(parseFloat(cookies.graph_token_expires - FIVE_MINUTES))
        if (expiration > new Date()) {
            return token;
        }
    }

    // Either no token or Expired Token.
    // Check whether have refresh_token

    const refresh_token = cookies.graph_refresh_token;

    if (refresh_token) {
        const newToken = await oauth2.accessToken.create({refresh_token: refresh_token}).refresh();
        saveValuesToCookie(newToken, res);
        return newToken.token.access_token;
    }

    // Nothing ?
    return null;
}

exports.getAccessToken = getAccessToken;
exports.clearCookies = clearCookies;
exports.getTokenFromCode = getTokenFromCode;
exports.getAuthUrl = getAuthUrl;