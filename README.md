# Running it (normal):
    `python3 intercept_min.py`

# Running tests:
    1. `python3 intercept_min.py --test self`

    2. osascript -e 'display notification "test" with title "Test"' or
        osascript -e 'display notification "Test entered the virtual fence" with title "🏠 Back In Safe Zone"'

    3. `cd safari-test; python3 -m http.server 8000`

    4. In safari: `http://localhost:8000`

        4.1 Allow notifications

        4.2 Send notifications via worker and plain


Run python3 notify_edit.py for the normal mode, where Tractive comes from the screen and everything else from the database.
Try python3 notify_edit.py --screen-all only if you want screen-only notifications beyond Tractive. Run it with --debug first, and watch for unexpected alerts being edited. I’m not sure what other texts appear in that window, such as when the Notification Center panel is open.

In --screen-all mode, any osascript test notification will be edited too, because the screen can’t tell which app posted it. The [Edited] title prefix still prevents loops.