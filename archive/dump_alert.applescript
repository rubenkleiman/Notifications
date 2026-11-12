-- Verbose probe: prints window names and, for every UI element, its role,
-- title, description and value. Errors are reported, not hidden.
-- Run while the Tractive alert is visible:  osascript dump_alert2.applescript

on clean(x)
	try
		set s to x as text
		if s is "missing value" then return ""
		return s
	on error
		return ""
	end try
end clean

tell application "System Events"
	tell process "NotificationCenter"
		set out to "Windows: " & (count of windows) & linefeed
		set wi to 0
		repeat with w in windows
			set wi to wi + 1
			set wname to ""
			try
				set wname to my clean(name of w)
			end try
			set out to out & "== Window " & wi & " name=[" & wname & "]" & linefeed
			try
				set els to entire contents of w
				set out to out & "   elements: " & (count of els) & linefeed
				repeat with e in els
					set r to ""
					set t to ""
					set d to ""
					set v to ""
					try
						set r to my clean(role of e)
					end try
					try
						set t to my clean(title of e)
					end try
					try
						set d to my clean(description of e)
					end try
					try
						set v to my clean(value of e)
					end try
					if t is not "" or d is not "" or v is not "" then
						set out to out & "   " & r & " | title=" & t & " | desc=" & d & " | value=" & v & linefeed
					end if
				end repeat
			on error errMsg
				set out to out & "   error reading contents: " & errMsg & linefeed
			end try
		end repeat
		return out
	end tell
end tell
