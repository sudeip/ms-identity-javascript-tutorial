const express = require('express');

const data = require('../controllers/data');

// initialize router
const router = express.Router();

router.get('/data', data.getData);

module.exports = router;
