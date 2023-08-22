# SharpSword
Developed By: @sadpanda_sec & @C0mmand3rOps3c

Description: 
   - Read Contents of Word Documents using MS Office Interop (Standalone or with CobaltStrike Execute Assembly)

Usage: 
   - SharpSword.exe C:\\Some\\Path\\To\\Document.(doc/docm/docx/etc...) [-checkPassword] -[password <password>]

Examples:

   - SharpSword.exe test.docx : read the contents of a word doc
   - SharpSword.exe test.docx -checkPassword : checks if the document is password protected
   - SharpSword.exe test.docx -password <somepassword> : decrypts the password protected document and reads contents in memory
