
Creates AD users

* First creates the user in English
* Later (if specified) updates the name to Hebrew
* At the end adds the user to the AD group (if specified)

Functionality:
* Checks if the username is available
* Checks if the password is complex (inaccurate)
* Removes unnecessary spaces and updates uppercase / lowercase letters
* After creation, verifies all the details:
    - User successfully created
    - User information updated
    - Found in all selected groups

In the "Command" panel you can see the command to be executed (PowerShell)


Planned:
* Option to add additional user information
* Save templates
* Copy details from an existing user


AutoHotkey Download Page:
https://autohotkey.com/download/

![alt text](https://github.com/benny779/ADUserCreator/blob/main/example.png?raw=true)
