
const { supabase } = require('./supabase');

const D = process.env.DATA_DIR || './data';

module.exports = { supabase, D };
