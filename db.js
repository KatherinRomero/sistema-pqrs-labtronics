const fs = require('fs'), Database = require('better-sqlite3');
const D = process.env.DATA_DIR || './data';
['files', 'pdf'].forEach(d => fs.mkdirSync(`${D}/${d}`, { recursive: true }));
const db = new Database(`${D}/pqrs.db`);
db.pragma('journal_mode = WAL');
db.exec("create table if not exists p(id text primary key,num integer,codigo text,ts text,pos integer default 0,fin integer default 0,d text default '{}')");
module.exports = { db, D };
