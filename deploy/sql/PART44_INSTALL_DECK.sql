-- PART 44 (4 Nov 2026): the Raise Your Exit Price Install deck is a document kind of its own.
-- Run once in the SQL editor. Adds 'exit_price_install' to the allowed document kinds.
do $$
declare def text;
begin
  select pg_get_constraintdef(oid) into def from pg_constraint where conname = 'documents_kind_check';
  if def is null then raise notice 'no documents_kind_check constraint found'; return; end if;
  if def like '%exit_price_install%' then raise notice 'already allowed'; return; end if;
  alter table documents drop constraint documents_kind_check;
  def := replace(def, '''valuation_builder''', '''valuation_builder'', ''exit_price_install''');
  execute 'alter table documents add constraint documents_kind_check ' || def;
end $$;
select pg_get_constraintdef(oid) from pg_constraint where conname = 'documents_kind_check';
