# GoPPT
Golang library for Powerpoint manipulation

This is not a complete openxml library. It does not convert xml into go objects.

What this code does is it will unzip ppt and replace some text placeholders and replace image place holder. Additionally it will duplicate one slide. Everything is done by regex replacements. This is not done in in-memory and the file system is windows. SO it may not work in other OS.
