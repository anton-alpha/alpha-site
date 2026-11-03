# Alpha waitlist backend

Google Apps Script bound to a Google Sheet in the Alpha Workspace. See the header of `Code.gs` for setup.
The website posts to the deployed web-app URL; confirmations are sent by email from the Workspace account that deployed it.

Sheet columns: id · timestamp · email · name · country · ref_by · code · confirmed · confirmed_at · token · ua · source

Position = order among confirmed signups, minus 50 places per confirmed referral.
