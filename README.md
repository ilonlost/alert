# FK Access Report

Dockerized morning/evening access-control reporting service for TimeShiftFK.

## What it does

- Morning report: events from 06:45 to 08:20, automatic email at 08:30.
- Evening report: events from 18:45 to 20:20, automatic email at 20:30.
- Counts unique employees by the first `ВХОД` through `ФК-1эт-ВЕСТИБЮЛЬ(ЛК9)` in the selected interval.
- Shows `Всего прошло` and one non-duplicated table of all departments.
- The department table is integrity-checked: its sum must equal `Всего прошло`.
- If a card/code is present in access events but missing from `GetFKPersonal`, it is counted under `Не найден в справочнике`, so totals never silently diverge.
- Attaches an XLSX file with employee code, full name, first passage time, department, full department path and position.
- Runs continuously in Docker and schedules both reports internally.

## Quick start on Linux

```bash
git clone https://github.com/ilonlost/alert.git
cd alert
cp .env.example .env
nano .env
```

At minimum set:

```env
DB_USER=...
DB_PASSWORD=...
MAIL_TO=user1@agrohold.ru,user2@agrohold.ru
```

If your SQL host name `11-vm-dwh01` is not resolvable from the Docker host, set `DB_SERVER` to the correct DNS name or IP.

Build:

```bash
docker compose build
```

## Safe Linux test — no email

Test the morning report without sending mail:

```bash
docker compose run --rm fk-access-report python app.py test-morning
```

Test the evening report without sending mail:

```bash
docker compose run --rm fk-access-report python app.py test-evening
```

The Excel file is written to `./reports/` on the Docker host.

Watch the console output and verify that:

```text
Unique employees: N; department sum: N
```

Both numbers must be equal.

## Test real email manually

Morning:

```bash
docker compose run --rm fk-access-report python app.py morning
```

Evening:

```bash
docker compose run --rm fk-access-report python app.py evening
```

These commands send mail.

## Start automatic mode

```bash
docker compose up -d --build
```

The container then stays running and sends automatically:

- 08:30 Europe/Moscow — morning report
- 20:30 Europe/Moscow — evening report

Check status:

```bash
docker compose ps
```

Watch logs:

```bash
docker compose logs -f
```

Restart:

```bash
docker compose restart
```

Stop:

```bash
docker compose down
```

Update from GitHub:

```bash
git pull
docker compose up -d --build
```

## Configuration

All passwords and recipients stay only in `.env`; `.env` is excluded from Git.

The Docker image uses the FreeTDS SQL Server ODBC driver from Debian, so it does
not need access to `packages.microsoft.com` during a build. Keep
`DB_DRIVER=FreeTDS` unless you intentionally install and select another driver.
For FreeTDS, `DB_PORT=1433` and `DB_TDS_VERSION=7.4` are the defaults.

Main variables:

```env
DB_SERVER=11-vm-dwh01
DB_NAME=TimeShiftFK
DB_USER=...
DB_PASSWORD=...

SMTP_HOST=smtp.agrohold.ru
SMTP_PORT=25
SMTP_USER=
SMTP_PASSWORD=
MAIL_FROM=WGSentry11@agrohold.ru
MAIL_TO=user1@agrohold.ru,user2@agrohold.ru

TARGET_AREA=ФК-1эт-ВЕСТИБЮЛЬ(ЛК9)
TIMEZONE=Europe/Moscow
```

If SMTP authentication is required later, fill `SMTP_USER` and `SMTP_PASSWORD`. Otherwise leave them empty for an internal relay.

## Notes

The evening mode currently uses the same event semantics as the working morning script: `ВХОД`, first passage per employee in the interval. If the evening report should mean employees leaving the site, change its rule to `ВЫХОД`/last exit before production use.
