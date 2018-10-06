var express = require('express');
var router = express.Router();
var authHelper = require('../helpers/auth');

/* GET home page. */
router.get('/', function(req, res, next) {
    let parms = { title: 'Home', active: { home: true } };

    parms.signInUrl = authHelper.getAuthUrl();
    parms.debug = parms.signInUrl;
    res.render('index', parms);
});

module.exports = router;