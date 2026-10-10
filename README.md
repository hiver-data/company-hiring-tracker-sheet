# Who is hiring? A Google Sheet that lists every open job at your companies

A Google Apps Script for a Google Sheet. Type company names or websites in column A, click `hiver > Fetch open jobs`, and the sheet fills with every open job from each company's own career board, plus a tab that shows which board each company matched and why. It uses the [Company Jobs Scraper: All Open Jobs by Company Name or Website](https://apify.com/hiver/all-open-jobs-by-company?utm_source=github&utm_campaign=company-hiring-tracker-sheet) Actor on Apify (Workday, Greenhouse, Lever, Ashby, SmartRecruiters and 12 more sources).

Useful for recruiters watching competitors, sales teams looking for companies that are hiring, and investors tracking headcount growth.

## Install in 5 steps

1. Make a new Google Sheet. Rename the first tab `Companies`. Put company names (`Spotify`) or websites (`stripe.com`) in column A, one per row. Websites match more reliably.
2. Open `Extensions > Apps Script`. Replace the contents of `Code.gs` with [apps-script/Code.gs](apps-script/Code.gs). (Optional: in Project Settings tick "Show appsscript.json" and paste [apps-script/appsscript.json](apps-script/appsscript.json) to limit the permissions the script asks for.) Save, then reload the sheet.
3. Create a free Apify account and copy your API token from Apify Console > Settings > API & Integrations. In the sheet use `hiver > Set Apify token`. The token is stored in Script Properties, never in a cell.
4. Optional: `hiver > Create Settings tab` adds filters (title, location, remote only, posted within N days) and the caps. Run it again after updating the script and it adds any new settings to your existing tab.
5. Click `hiver > Fetch open jobs` and approve the permissions on the first run. The `Jobs` tab fills in; the `Matches` tab lists the career board used for each company, the number of jobs written for it, and the evidence, or why a company was not matched.

## Daily "what is new" tab

`hiver > Turn on daily new-jobs check` adds a daily trigger that runs the Actor with `onlyNewSinceLastRun: true`. The first run delivers every open job; later runs append only jobs not delivered before to a `New jobs` tab, with the date found in column A. You are charged for new jobs only.

## How it works

- It starts an asynchronous Apify run (`POST /v2/acts/hiver~all-open-jobs-by-company/runs`) and polls it, then reads the dataset. It does not use the synchronous endpoint, because Apps Script's `UrlFetchApp` gives up on long requests: on the Lab bench a run took 17 to 98 seconds for 2 to 5 companies, and longer lists take longer.
- Apps Script stops a script after 6 minutes. The script polls for at most 4.5 minutes, then sets a one-minute trigger and continues from the stored run id. `hiver > Check / resume running fetch` does the same by hand.
- `Jobs` columns are the same as in the [CSV examples](https://github.com/hiver-data/all-open-jobs-by-company): company, source, title, location, workplaceType, isRemote, department, employmentType, postedAt, salaryMin, salaryMax, salaryCurrency, salaryInterval, url, applyUrl, inputCompany.

## What it costs

The Actor charges $1.50 per 1,000 jobs on the Apify Free and Starter plans (less on Scale and Business), plus Apify's $0.00005 start fee per run. Companies that are not found or have no open jobs cost nothing. By default a run delivers at most 500 jobs and stops once it has cost $1.00 (`maxJobs` and `maxChargeUsd` in the Settings tab). `maxJobs` counts across all companies, so one large employer near the top of the list can use it up before the companies below it are read: `maxJobsPerCompany` (default 25 in the Settings tab) shares it out so every company gets rows; keep `maxJobs` at least companies x `maxJobsPerCompany`, or the sheet's start message warns that later companies may get fewer rows (set it to 0 for no per-company cap). When a run stops at `maxJobs`, the sheet says so and the `Matches` tab shows which companies got 0 rows. The free Apify plan includes $5 of usage a month.

## What we measured

Lab bench, 2026-10-10 08:16 UTC, Actor build 0.3.22, four test cases with the same inputs for each Actor. Each answer was checked against the company's real board. These are the Actor's numbers, not the sheet's.

| | This Actor | glitchbound/ats-jobs-scraper | fabro.dev/company-career-page-jobs |
|---|---|---|---|
| Accuracy (job URL on the company's real board, system named correctly, pay range) | 100% | 68% | 23% |
| Completeness (job fields filled) | 88% | 58% | 57% |
| Test cases that returned the companies' jobs (at least the bench minimum) | 4 of 4 | 3 of 4 | 1 of 4 |
| Runtime per test case (average of 4) | 39.5 s | 19.3 s | 3.3 s |
| Price per 1,000 jobs in these runs | $1.51 | $2.34 | $2.01 |

The other two are faster. Four cases and a handful of companies is a small sample; the bench runs again after every new build.

## Honest limits

- A company is matched only when its own website links to the board, or the board or its job ads name the company. Otherwise it shows as `not matched` instead of a guess ("Tide" is a bank and a detergent). Large employers with in-house career sites that publish no public job feed may not match at all.
- A sheet holds up to 10 million cells; very large result sets (many thousands of jobs) are better downloaded from Apify as CSV.
- Job descriptions are not written to the sheet (cells are limited to 50,000 characters and it keeps the sheet light).
- The script is not published as a Google add-on; you paste it into your own sheet and Google asks you to approve it.

## Tests

```
npm test
```

Runs `Code.gs` in Node with mocked `UrlFetchApp`, `SpreadsheetApp`, `PropertiesService` and `ScriptApp`, over job rows recorded from a real board run of the Actor (`test/fixtures`). It makes no network calls and no Apify run. Node 18+.

Live check, 2026-10-10: the previous version of `Code.gs` (commit a1d4b64) ran unchanged against the real Apify API, with Google's Sheets, properties and triggers simulated: 2 companies (stripe.com, NVIDIA), run SUCCEEDED in 21 seconds, 10 jobs written, both companies matched. All 10 rows were Stripe's because `maxJobs` was 10; that run is why `maxJobsPerCompany` and the jobs-per-company column exist. The script has not yet been run inside a real Google Sheet.

## Notes

- Unofficial. Not affiliated with Google, Apify, any job system named here or any company whose jobs you read. The Actor reads only public job-board data that companies publish for job seekers, with no login.
- Something wrong or missing for a company? Open an issue here or on the Actor's Issues tab on Apify.

MIT licensed. Built by [hiver](https://apify.com/hiver?utm_source=github&utm_campaign=company-hiring-tracker-sheet).
