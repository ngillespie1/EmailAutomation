# EmailAutomation

These are files to assist in the automation of extracting and normalizing emails in outlook, to use for machine learning, as well as aiding productivity by using folders to trigger actions.

In order:
1. PullAttachmentsOutOfMSGsInFolder: This is designed to take the MSG's that are contained within a folder and extract their attachments to disk. In my use case this was because the MSGs of interest were themselves attachments, but this could be repurposed to just extract a bunch of attachments.
2. ChunkMSGintoSingleLineCSV: This takes the extracted MSGs from (1) above, and parses them into a single line text file, to be used to train the ML Model.
3. MLCategorizationOfEmail: This script uses Keras and Tensorflow, and a prepopulated ML model to read mails within a folder within Outlook, and take certain actions. EG if the model believes the mail can be categorized as A, it will perform a certain serious of actions, similarly with category B it will take an alternate set of actions.
4. AutomatedActionAfterPlacingEmailinOutlookFolder: Does what it says on the tin. If the email is placed in a folder, it will take action, including sending a mail with predesignated attachment.
