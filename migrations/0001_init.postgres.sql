-- wOS Sheets tables (Postgres). Every table starts with sheets_ and has team_id, per the suite contract,
-- so Sheets can share a database with other wOS apps. Ids, times and JSON are text.
create table if not exists sheets_teams (id text primary key, name text not null, created_at text not null);
create table if not exists sheets_people (
  id text primary key, team_id text not null, name text not null, email text, github text, account_sub text,
  role text not null default 'member', prefs text, created_at text not null, deactivated_at text);
create index if not exists sheets_people_team on sheets_people (team_id);
create table if not exists sheets_workbooks (
  id text primary key, team_id text not null, title text not null, created_by text, created_at text not null,
  updated_at text not null, updated_by text, state text, state_seq bigint not null default 0,
  share_token text, share_mode text not null default 'off', example integer not null default 0, deleted_at text);
create index if not exists sheets_workbooks_team on sheets_workbooks (team_id, updated_at);
create unique index if not exists sheets_workbooks_share on sheets_workbooks (share_token) where share_token is not null;
create table if not exists sheets_updates (id bigserial primary key, team_id text not null, workbook_id text not null, data text not null, by_id text, created_at text not null);
create index if not exists sheets_updates_book on sheets_updates (workbook_id, id);
create table if not exists sheets_versions (
  id text primary key, team_id text not null, workbook_id text not null, label text, auto integer not null default 0,
  state text not null, created_by text, created_at text not null);
create index if not exists sheets_versions_book on sheets_versions (workbook_id, created_at);
create table if not exists sheets_files (
  id text primary key, team_id text not null, uploader_id text, name text not null, type text not null, size integer not null,
  data text not null, created_at text not null);
create table if not exists sheets_approvals (
  id text primary key, team_id text not null, person_id text not null, requested_by text not null, tool text not null, input text not null,
  status text not null default 'waiting', result text, created_at text not null, decided_at text);
create table if not exists sheets_settings (team_id text not null, key text not null, value text not null, primary key (team_id, key));
