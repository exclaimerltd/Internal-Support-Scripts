#!/usr/bin/env /usr/bin/sudo /bin/bash

if [ -f /Library/LaunchAgents/com.exclaimer.csua.plist ]; then
    for U in $(who | awk '/console/ { print $1 }')
    do
        launchctl bootout gui/$(id -u $U) /Library/LaunchAgents/com.exclaimer.csua.plist
    done
fi

rm -rf /Applications/Exclaimer\ Signature\ Agent.app \
            /Library/Exclaimer/Exclaimer\ Signature\ Agent.app \
            /Library/LaunchAgents/com.exclaimer.csua.plist

pkgutil --forget com.exclaimer.csua
