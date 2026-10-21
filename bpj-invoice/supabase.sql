-- Run once in Supabase SQL Editor. All data access is through the server-side service key.
create table if not exists public.invoice_counter (id integer primary key check (id=1), last_no integer not null);
insert into public.invoice_counter(id,last_no) values (1,20) on conflict (id) do nothing;
create table if not exists public.invoices (
 id uuid primary key default gen_random_uuid(),
 invoice_no integer unique not null,
 invoice_date date not null,
 template text not null,
 customer text not null,
 items jsonb not null,
 total numeric(12,2) not null,
 created_at timestamptz not null default now()
);
create index if not exists invoices_date_idx on public.invoices(invoice_date);
create or replace function public.create_invoice(p_date date,p_template text,p_customer text,p_items jsonb,p_total numeric)
returns public.invoices language plpgsql security definer set search_path=public as $$
declare n integer; rec public.invoices;
begin
 update public.invoice_counter set last_no=last_no+1 where id=1 returning last_no into n;
 if n is null then raise exception 'Counter missing'; end if;
 insert into public.invoices(invoice_no,invoice_date,template,customer,items,total)
 values(n,p_date,p_template,p_customer,p_items,p_total) returning * into rec;
 return rec;
end $$;
revoke all on function public.create_invoice(date,text,text,jsonb,numeric) from public,anon,authenticated;
alter table public.invoices enable row level security;
alter table public.invoice_counter enable row level security;
-- No client RLS policies: service-role backend only.
