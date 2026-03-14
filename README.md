# Phone App

1. Copy your Supabase project URL and anon key into `supabase-config.js`.
2. Run `supabase-setup.sql` in the Supabase SQL editor.
3. Reload the app, open the Data section, and use Cloud Backup to sign up or sign in.
4. Once signed in, the app uploads a private backup and can restore it after local browser data is cleared.

Notes:
- If Supabase email confirmation is enabled, confirm the email before signing in.
- Clearing browser data still signs the device out, but the cloud backup remains in Supabase.

