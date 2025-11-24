Here’s a brief explanation of what our code does:

Upload & Extract Emails:
The code lets you select a ZIP file containing .eml emails.
It automatically extracts all .eml files into a folder.

Parse Emails:
Each email is read and relevant information is extracted: sender, recipient, subject, date, message ID, and the body of the email.
The body is cleaned: it removes signatures, greetings, disclaimers, quoted content, URLs, and punctuation, and converts text to lowercase.

Build Email Threads:
Emails are connected based on In-Reply-To and References headers to understand which email replies to which.
A thread graph is built internally (using Python’s networkx) showing the relationships.

Export to CSV:
All parsed emails are exported into a CSV file for easy review. Each row corresponds to a single email with its cleaned content and metadata.
