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

async function getTokenFromCode(auth_code) {
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
    console.log("HERE");
    console.log('Token Created', token);
    return token.token.access_token;
}

exports.getTokenFromCode = getTokenFromCode;
exports.getAuthUrl = getAuthUrl;