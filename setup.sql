-- Supabase v0.1 compatibility migration
-- adds JSON state to rooms table. Existing tables remain.
alter table public.rooms add column if not exists game_state jsonb;

-- atomic pick submission: prevents concurrent users overwriting each other's pick entries
create or replace function public.dbb_submit_pick(p_room_code text, p_user_id text, p_pick jsonb)
returns void language plpgsql security invoker as $$
begin
  update public.rooms
  set game_state = jsonb_set(
    coalesce(game_state, '{}'::jsonb),
    '{picks}',
    coalesce(game_state->'picks', '{}'::jsonb) || jsonb_build_object(p_user_id, p_pick),
    true
  )
  where room_code = p_room_code;
  if not found then raise exception 'room not found'; end if;
end;
$$;

grant execute on function public.dbb_submit_pick(text, text, jsonb) to anon, authenticated;

-- ensure realtime publication has rooms table (safe if already added)
do $$ begin
 if not exists (select 1 from pg_publication_tables where pubname='supabase_realtime' and schemaname='public' and tablename='rooms') then
  alter publication supabase_realtime add table public.rooms;
 end if;
end $$;
