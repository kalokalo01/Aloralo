# buyerbotv1

## Run
1. Install dependencies: `pip install -r requirements.txt`
2. Set `BOT_TOKEN` and `ADMIN_ID` as environment variables if using a version of the bot that reads them from the environment.
3. Start: `python buyerbotv1.py`

### Railway
The included `Procfile` runs the bot as a worker.

The bot uses SQLite (`bot.db`) and creates the `uploads/` folder automatically.
