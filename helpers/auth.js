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
}

exports.getTokenFromCode = getTokenFromCode;
exports.getAuthUrl = getAuthUrl;