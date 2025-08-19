Currently, it is set up to send a notification email to abc@gmail.com. The email is hard coded so just change it to whatever email you want it to send to.
The columnIndexes are the ones I'm specifically checking for, change it to whatever you are looking for.
The expirationRules are what I have set it to filter for.
When handling edgecases, I'm skipping the cells that don't have a date or have 'missing' or 'discharged' written on them. Can also change this.
