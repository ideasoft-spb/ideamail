import easygmail
from easygmail import mailbox

gmail = easygmail.Gmail()
search = mailbox.Search(gmail, max_results=5, labels=['UNREAD'])
print(search.full_results()[1].html)
