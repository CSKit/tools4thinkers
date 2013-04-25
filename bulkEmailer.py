#! /usr/bin/env python

import sys
import xlrd
from smtplib import SMTP
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def prepare_message_for(first, last, category):
  if category == 'Academic':
		message = """<html><body><p>Dear """+first+""" """+last+""",</p>
		<p>The Christian Science Organization at Cal would like to invite you and your students to a free event next Monday, April 22 at 8 pm. We are hosting a Christian Science practitioner (someone who prays for and heals people for a living) to talk about: 
			<ul><li>The nature of spiritual reality</li>
			<li>What spiritual reality means for us in our everyday lives</li>
			<li>How we pray</li></ul>
		It will be followed by an open discussion in small groups about questions like applying practical prayer in daily life or what the true nature of reality is.</p>
		<p>We are hoping to have a diverse group of spiritual seekers and critical thinkers. If any of your students fit this description, or just want an opportunity to see what Christian Science is all about, please send them to us! We will be open for drop-ins and chatting from 7 pm until late, if you can't make the regular lecture time please feel free to stop by anyway for a bit and check us out.</p>
		<hr/><p><b><i>Lecture and Chat: What is Spiritual Reality?</b></i></p>
		<p><b>Monday April 22 </b>at <b>8 pm</b>
			<br/>Christian Science Organization Building
			<br/><b>2601 Durant Ave, Berkeley</b>
			<br/><i>Free + refreshments</i>
		</p><hr/>
		<p>Regards,<br/>Members of the Christian Science Organization</p>
		</body></html>"""
	if category == 'CS':
		message = """<html><body><p>Dear """+first+""",</p>
		<p>The Christian Science Organization at Cal Berkeley is hosting a lecture next Monday. We are hoping to have anyone from your membership or attendees join us and would appreciate your help in spreading the word in whatever ways seem appropriate.</p>
		<p>In particular, we would love to have strong attendance from students in high school and college and recent graduates, since we are reaching out to the campus community. Would you please share the details of the lecture with your Sunday School superintendent so that they can help us reach out to that audience? We would be more than happy to help arrange carpools and other transportation.</p>
		<hr/><p><b><i>Lecture and Chat: What is Spiritual Reality?</b></i></p>
		<p><b>Monday April 22 </b>at <b>8 pm</b>
			<br/>Christian Science Organization Building
			<br/><b>2601 Durant Ave, Berkeley</b>
			<br/><i>Free</i>
		</p><hr/>
		<p>Love,<br/>Members of the Christian Science Organization</p>
		</body></html>"""
	return message

def prepare_plaintext_message_for(first, last, category):
	if category == 'Academic':
		message = """Dear """+first+""" """+last+""",\n\nThe Christian Science Organization at Cal would like to invite you and your students to a free event next Monday, April 22 at 8 pm. We are hosting a Christian Science practitioner (someone who prays for and heals people for a living) to talk about:\n* The nature of spiritual reality\n* What spiritual reality means for us in our everyday lives\n* How we pray\nIt will be followed by an open discussion in small groups about questions like applying practical prayer in daily life or what the true nature of reality is.\n\nWe are hoping to have a diverse group of spiritual seekers and critical thinkers. If any of your students fit this description, or just want an opportunity to see what Christian Science is all about, please send them to us! We will be open for drop-ins and chatting from 7 pm until late, if you can't make the regular lecture time please feel free to stop by anyway for a bit and check us out.\n\n____________________________\nLecture and Chat: What is Spiritual Reality?\n\nMonday April 22 at 8 pm\nChristian Science Organization Building\n2601 Durant Ave, Berkeley\nFree + refreshments\n_____________________________\nRegards,\nMembers of the Christian Science Organization"""
	if category == 'CS':
		message = """Dear """+first+""",\n\nThe Christian Science Organization at Cal Berkeley is hosting a lecture next Monday. We are hoping to have anyone from your membership or attendees join us and would appreciate your help in spreading the word in whatever ways seem appropriate.\n\nIn particular, we would love to have strong attendance from students in high school and college and recent graduates, since we are reaching out to the campus community. Would you please share the details of the lecture with your Sunday School superintendent so that they can help us reach out to that audience? We would be more than happy to help arrange carpools and other transportation.\n\n____________________________\nLecture and Chat: What is Spiritual Reality?\n\nMonday April 22 at 8 pm\nChristian Science Organization Building\n2601 Durant Ave, Berkeley\nFree\n_____________________________\n\nLove,\nMembers of the Christian Science Organization"""
	return message

def get_send_data():
	wb = xlrd.open_workbook('/Users/noah/Documents/Personal/cso_send_list.xlsx')
	sh = wb.sheet_by_index(0)
	email_list = []
	for rownum in range(sh.nrows):
		if rownum > 0:
			email_list.append(sh.row_values(rownum))
	return email_list

def send_message(password, sender, recipient, message, message_plaintext):
	# Create message container - the correct MIME type is multipart/alternative.
	msg = MIMEMultipart('alternative')
	msg['Subject'] = "Cal Christian Science Organization Community Event, Monday April 22, 8 PM"
	msg['From'] = sender
	msg['To'] = recipient

	# Create the body of the message (a plain-text and an HTML version).
	text = message_plaintext
	html = message

	# Record the MIME types of both parts - text/plain and text/html.
	part1 = MIMEText(text, 'plain')
	part2 = MIMEText(html, 'html')

	# Attach parts into message container.
	# According to RFC 2046, the last part of a multipart message, in this case
	# the HTML message, is best and preferred.
	msg.attach(part1)
	msg.attach(part2)

	# Send the message via gmail SMTP server.
	server = SMTP('smtp.gmail.com:587')
	server.ehlo()
	server.starttls()
	server.login(sender, password)
	# sendmail function takes 3 arguments: sender's address, recipient's address
	# and message to send - here it is sent as one string.
	try:
		server.sendmail(sender, recipient, msg.as_string())
	finally:
		server.quit()

def main(argv):
	sender = 'csoberkeley@gmail.com'
	send_data = get_send_data()
	password = argv[1]
	for row in send_data:
		recipient = row[2]
		first = row[0]
		last = row[1]
		category = row[3]
		message = prepare_message_for(first, last, category)
		message_plaintext = prepare_plaintext_message_for(first, last, category)
		send_message(password, sender, recipient, message, message_plaintext)

if __name__ == "__main__":
	main(sys.argv)
