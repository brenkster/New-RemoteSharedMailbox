# New-RemoteSharedMailbox
Create a shared mailbox from hybrid Exchange directly in Exchange Online but also visible in on-premises Active directory

How to create a shared mailbox from hybrid Exchange directly in Exchange Online but also visible in on-premises Active directory
Now that's a mouthful, but a question that's asked often.
And that's not strange at all, because after moving most of your mailboxes to Exchange Online, the inevitable new shared mailbox request will pop up.

Only to find out that if created in the old way you need to manually move it Exchange Online.

And if you create shared mailbox in the ECP you will soon find that you problably made the wrong choice and the new mailbox is not visible in your on-premises Active Directory.

There's where this script comes to the recue.
It creates a remote shared mailbox, a distribution group for full access and send-as rights, adds that group to the mailbox, hides the distribution group because we only use it for access rights.
Then login to Exchange Online, disable POP, IMAP, Activesync and OWA, and resets the proxy settings if neccesary.

This can also be used to create usermailboxes, just tweak it to your needs.
