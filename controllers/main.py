#!/usr/bin/env python
#
# Copyright 2007 Google Inc.
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.
#
"""Starting template for Google App Engine applications.

Use this project as a starting point if you are just beginning to build a Google
App Engine project. Remember to download the OAuth 2.0 client secrets which can
be obtained from the Developer Console <https://code.google.com/apis/console/>
and save them as 'client_secrets.json' in the project directory.
"""

__author__ = 'jcgregorio@google.com (Joe Gregorio)'


import httplib2
import logging
import os
import pickle
import webapp2

from controllers.base import BaseHandler
from apiclient.discovery import build
from oauth2client.appengine import oauth2decorator_from_clientsecrets
from oauth2client.client import AccessTokenRefreshError
from google.appengine.api import memcache
from apiclient import errors

# CLIENT_SECRETS, name of a file containing the OAuth 2.0 information for this
# application, including client_id and client_secret, which are found
# on the API Access tab on the Google APIs
# Console <http://code.google.com/apis/console>
CLIENT_SECRETS = os.path.join(os.path.dirname(os.path.dirname(__file__)),
    'client_secrets.json')

# Helpful message to display in the browser if the CLIENT_SECRETS file
# is missing.
MISSING_CLIENT_SECRETS_MESSAGE = """
<h1>Warning: Please configure OAuth 2.0</h1>
<p>
To make this sample run you will need to populate the client_secrets.json file
found at:
</p>
<p>
<code>%s</code>.
</p>
<p>with information found on the <a
href="https://code.google.com/apis/console">APIs Console</a>.
</p>
""" % CLIENT_SECRETS


http = httplib2.Http(memcache)
service = build("plus", "v1", http=http)
decorator = oauth2decorator_from_clientsecrets(
    CLIENT_SECRETS,
    ['https://www.googleapis.com/auth/plus.me',
    'https://www.googleapis.com/auth/gmail.modify',
    'https://www.googleapis.com/auth/gmail.labels',
    'https://www.googleapis.com/auth/drive.readonly',
    'https://spreadsheets.google.com/feeds',
    'https://www.googleapis.com/auth/gmail.labels'],
    MISSING_CLIENT_SECRETS_MESSAGE)

def ListThreadsMatchingQuery(service, user_id, query=''):
  """List all Threads of the user's mailbox matching the query.

  Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    query: String used to filter messages returned.
           Eg.- 'label:UNREAD' for unread messages only.

  Returns:
    List of threads that match the criteria of the query. Note that the returned
    list contains Thread IDs, you must use get with the appropriate
    ID to get the details for a Thread.
  """
  try:
    response = service.users().threads().list(userId=user_id, q=query).execute()
    threads = []
    if 'threads' in response:
      threads.extend(response['threads'])

    while 'nextPageToken' in response:
      page_token = response['nextPageToken']
      response = service.users().threads().list(userId=user_id, q=query,
                                        pageToken=page_token).execute()
      threads.extend(response['threads'])

    return threads
  except errors.HttpError, error:
    print 'An error occurred: %s' % error

def ListLabels(service, user_id):
  """Get a list all labels in the user's mailbox.

  Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.

  Returns:
    A list all Labels in the user's mailbox.
  """
  try:
    response = service.users().labels().list(userId=user_id).execute()
    labels = response['labels']
    for label in labels:
      print 'Label id: %s - Label name: %s' % (label['id'], label['name'])
    return labels
  except errors.HttpError, error:
    print 'An error occurred: %s' % error

def CreateLabel(service, user_id, label_object):
  """Creates a new label within user's mailbox, also prints Label ID.

  Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    label_object: label to be added.

  Returns:
    Created Label.
  """
  try:
    label = service.users().labels().create(userId=user_id,
                                            body=label_object).execute()
    print label['id']
    return label
  except errors.HttpError, error:
    print 'An error occurred: %s' % error

def MakeLabel(label_name, mlv='show', llv='labelShow'):
  """Create Label object.

  Args:
    label_name: The name of the Label.
    mlv: Message list visibility, show/hide.
    llv: Label list visibility, labelShow/labelHide.

  Returns:
    Created Label.
  """
  label = {'messageListVisibility': mlv,
           'name': label_name,
           'labelListVisibility': llv}
  return label

def ModifyThread(service, user_id, thread_id, msg_labels):
  """Add labels to a Thread.

  Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    thread_id: The id of the thread to be modified.
    msg_labels: The change in labels.

  Returns:
    Thread with modified Labels.
  """
  try:
    thread = service.users().threads().modify(userId=user_id, id=thread_id,
                                              body=msg_labels).execute()

    thread_id = thread['id']
    label_ids = thread['messages'][0]['labelIds']

    print 'Thread ID: %s - With Label IDs %s' % (thread_id, label_ids)
    return thread
  except errors.HttpError, error:
    print 'An error occurred: %s' % error


def CreateMsgLabels(label_list):
  """Create object to update labels.

  Returns:
    A label update object.
  """
  return {'removeLabelIds': [], 'addLabelIds': label_list }

class MainHandler(BaseHandler):

  @decorator.oauth_aware
  def get(self):
    variables = {
        'url': decorator.authorize_url(),
        'has_credentials': decorator.has_credentials()
        }
    self.render_response('grant.html', **variables)

class AboutHandler(BaseHandler):

  @decorator.oauth_required
  def get(self):
    try:
      http = decorator.http()
      user = service.people().get(userId='me').execute(http)
      text = 'Hello, %s!' % user['displayName']



      import gspread
      credentials = decorator.credentials
      gc = gspread.authorize(credentials)
      sh = gc.open_by_key('1u08koEA4i4vZxJP9P3vd55pz48mzDOJ72IN57QnIjug')
      worksheet = sh.get_worksheet(0)
      emails_3termo = worksheet.col_values(2)

      sh = gc.open_by_key('1f6nAsU5nLiKpzh0X7zAB3_vuEPk0Vs9m6jDQ02c6yu4')
      emails_5termo = sh.get_worksheet(0).col_values(2)


      gmail_service = build('gmail', 'v1', http=http)


      def get_label_or_create(labels, label_name):
        for label in labels:
          if label['name'] == label_name:
            return label
        label_object = MakeLabel(label_name = label_name)
        label = CreateLabel(service=gmail_service, user_id='me', label_object = label_object )
        return label

      labels = ListLabels(service=gmail_service, user_id = 'me')

      label_aluno = get_label_or_create(labels = labels, label_name = 'ALUNO')
      label_3termo = get_label_or_create(labels = labels, label_name = '3Termo_2015')
      label_5termo = get_label_or_create(labels = labels, label_name = '5Termo_2015')


      def apply_label_in_students(students_email_list, label_list):
        data = {}
        for email in students_email_list:
          threads = ListThreadsMatchingQuery(service = gmail_service, user_id = 'me', query='from:{}'.format(email) )
          message_labes = CreateMsgLabels(label_list = label_list)
          for thread in threads:
            thread_id = thread['id']
            ModifyThread(service = gmail_service, user_id = 'me', thread_id = thread_id, msg_labels = message_labes)

      apply_label_in_students(students_email_list =  emails_3termo, label_list = [label_aluno['id'], label_3termo['id']])
  
      apply_label_in_students(students_email_list =  emails_5termo, label_list = [label_aluno['id'], label_5termo['id']])


      self.render_response('welcome.html', text=text, user=user
        , emails_3termo = emails_3termo, data = data, labels = labels, label_aluno = label_aluno)
    except AccessTokenRefreshError:
      self.redirect('/')




app = webapp2.WSGIApplication(
    [
     ('/', MainHandler),
     ('/about', AboutHandler),
     (decorator.callback_path, decorator.callback_handler())
    ],
    debug=True)
