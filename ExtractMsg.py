#!/usr/bin/env python
# -*- coding: latin-1 -*-
"""
ExtractMsg:
    Extracts emails and attachments saved in Microsoft Outlook's .msg files

https://github.com/mattgwwalker/msg-extractor
"""

__author__ = "Matthew Walker"
__date__ = "2013-11-19"
__version__ = '0.2'

# --- LICENSE -----------------------------------------------------------------
#
#    Copyright 2013 Matthew Walker
#
#    This program is free software: you can redistribute it and/or modify
#    it under the terms of the GNU General Public License as published by
#    the Free Software Foundation, either version 3 of the License, or
#    (at your option) any later version.
#
#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU General Public License for more details.
#
#    You should have received a copy of the GNU General Public License
#    along with this program.  If not, see <http://www.gnu.org/licenses/>.

import os
import sys
import glob
import traceback
import hashlib
from email.parser import Parser as EmailParser
import email.utils
import olefile as OleFile
import json
from imapclient import  imapclient

# This property information was sourced from
# http://www.fileformat.info/format/outlookmsg/index.htm
# on 2013-07-22.
properties = {
    '001A': 'Message class',
    '0037': 'Subject',
    '003D': 'Subject prefix',
    '0040': 'Received by name',
    '0042': 'Sent repr name',
    '0044': 'Rcvd repr name',
    '004D': 'Org author name',
    '0050': 'Reply rcipnt names',
    '005A': 'Org sender name',
    '0064': 'Sent repr adrtype',
    '0065': 'Sent repr email',
    '0070': 'Topic',
    '0075': 'Rcvd by adrtype',
    '0076': 'Rcvd by email',
    '0077': 'Repr adrtype',
    '0078': 'Repr email',
    '007d': 'Message header',
    '0C1A': 'Sender name',
    '0C1E': 'Sender adr type',
    '0C1F': 'Sender email',
    '0E02': 'Display BCC',
    '0E03': 'Display CC',
    '0E04': 'Display To',
    '0E1D': 'Subject (normalized)',
    '0E28': 'Recvd account1 (uncertain)',
    '0E29': 'Recvd account2 (uncertain)',
    '1000': 'Message body',
    '1008': 'RTF sync body tag',
    '1035': 'Message ID (uncertain)',
    '1046': 'Sender email (uncertain)',
    '3001': 'Display name',
    '3002': 'Address type',
    '3003': 'Email address',
    '39FE': '7-bit email (uncertain)',
    '39FF': '7-bit display name',

    # Attachments (37xx)
    '3701': 'Attachment data',
    '3703': 'Attachment extension',
    '3704': 'Attachment short filename',
    '3707': 'Attachment long filename',
    '370E': 'Attachment mime tag',
    '3712': 'Attachment ID (uncertain)',

    # Address book (3Axx):
    '3A00': 'Account',
    '3A02': 'Callback phone no',
    '3A05': 'Generation',
    '3A06': 'Given name',
    '3A08': 'Business phone',
    '3A09': 'Home phone',
    '3A0A': 'Initials',
    '3A0B': 'Keyword',
    '3A0C': 'Language',
    '3A0D': 'Location',
    '3A11': 'Surname',
    '3A15': 'Postal address',
    '3A16': 'Company name',
    '3A17': 'Title',
    '3A18': 'Department',
    '3A19': 'Office location',
    '3A1A': 'Primary phone',
    '3A1B': 'Business phone 2',
    '3A1C': 'Mobile phone',
    '3A1D': 'Radio phone no',
    '3A1E': 'Car phone no',
    '3A1F': 'Other phone',
    '3A20': 'Transmit dispname',
    '3A21': 'Pager',
    '3A22': 'User certificate',
    '3A23': 'Primary Fax',
    '3A24': 'Business Fax',
    '3A25': 'Home Fax',
    '3A26': 'Country',
    '3A27': 'Locality',
    '3A28': 'State/Province',
    '3A29': 'Street address',
    '3A2A': 'Postal Code',
    '3A2B': 'Post Office Box',
    '3A2C': 'Telex',
    '3A2D': 'ISDN',
    '3A2E': 'Assistant phone',
    '3A2F': 'Home phone 2',
    '3A30': 'Assistant',
    '3A44': 'Middle name',
    '3A45': 'Dispname prefix',
    '3A46': 'Profession',
    '3A48': 'Spouse name',
    '3A4B': 'TTYTTD radio phone',
    '3A4C': 'FTP site',
    '3A4E': 'Manager name',
    '3A4F': 'Nickname',
    '3A51': 'Business homepage',
    '3A57': 'Company main phone',
    '3A58': 'Childrens names',
    '3A59': 'Home City',
    '3A5A': 'Home Country',
    '3A5B': 'Home Postal Code',
    '3A5C': 'Home State/Provnce',
    '3A5D': 'Home Street',
    '3A5F': 'Other adr City',
    '3A60': 'Other adr Country',
    '3A61': 'Other adr PostCode',
    '3A62': 'Other adr Province',
    '3A63': 'Other adr Street',
    '3A64': 'Other adr PO box',

    '3FF7': 'Server (uncertain)',
    '3FF8': 'Creator1 (uncertain)',
    '3FFA': 'Creator2 (uncertain)',
    '3FFC': 'To email (uncertain)',
    '403D': 'To adrtype (uncertain)',
    '403E': 'To email (uncertain)',
    '5FF6': 'To (uncertain)'}


def windowsUnicode(string):
    if string is None:
        return None
    if sys.version_info[0] >= 3:  # Python 3
        return str(string, 'utf_16_le')
    else:  # Python 2
        return unicode(string, 'utf_16_le')



class Attachment:
    def __init__(self, msg, dir_):
        # Get long filename
        self.longFilename = msg._getStringStream([dir_, '__substg1.0_3707'])

        # Get short filename
        self.shortFilename = msg._getStringStream([dir_, '__substg1.0_3704'])

        # Get attachment data
        self.data = msg._getStream([dir_, '__substg1.0_37010102'])

    @property
    def sha1(self):
        try:
            return self._sha1
        except Exception:
            sha1 = hashlib.sha1()
            sha1.update(self.data)
            self._sha1 = sha1.hexdigest()
            return self._sha1

    def save(self):
        # Use long filename as first preference
        filename = self.longFilename
        # Otherwise use the short filename
        if filename is None:
            filename = self.shortFilename
        # Otherwise just make something up!
        if filename is None:
            import random
            import string
            filename = 'UnknownFilename ' + \
                ''.join(random.choice(string.ascii_uppercase + string.digits)
                        for _ in range(5)) + ".bin"
        f = open(filename, 'wb')
        f.write(self.data)
        f.close()
        return filename


    def toJson(self):
        return {
            'short_name': self.shortFilename,
            'long_name': self.longFilename,
        }

class OLEMessage(OleFile.OleFileIO):
    def __init__(self, filename):
        OleFile.OleFileIO.__init__(self, filename)

    def getStream(self, filename):
        if self.exists(filename):
            stream = self.openstream(filename)
            return stream.read()
        else:
            return None

    def getStringStream(self, filename, prefer='unicode'):
        """Gets a string representation of the requested filename.
        Checks for both ASCII and Unicode representations and returns
        a value if possible.  If there are both ASCII and Unicode
        versions, then the parameter /prefer/ specifies which will be
        returned.
        """

        if isinstance(filename, list):
            # Join with slashes to make it easier to append the type
            filename = "/".join(filename)

        asciiVersion = self.getStream(filename + '001E')
        unicodeVersion = windowsUnicode(self.getStream(filename + '001F'))
        if asciiVersion is None:
            return unicodeVersion
        elif unicodeVersion is None:
            return asciiVersion
        else:
            if prefer == 'unicode':
                return unicodeVersion
            else:
                return asciiVersion

class Message():
    def __init__(self, msgFilePath = None, oleMessage = None,root_path=[]):
        if msgFilePath is None and oleMessage is None:
            raise Exception("No message specified")
        if (not msgFilePath is None) and (not oleMessage is None):
            raise Exception("Use either a file path or OLEMessage object")
        self.root_path = root_path
        if msgFilePath is not None:
            self.oleMessage = OLEMessage(msgFilePath)
        else:
            self.oleMessage = oleMessage

    def _getStringStream(self, streamPath):
        if type(streamPath) is list:
            return self.oleMessage.getStringStream(self.root_path + streamPath)
        else:
            return self.oleMessage.getStringStream(self.root_path + [streamPath])

    def _getStream(self, streamPath):
        if type(streamPath) is list:
            return self.oleMessage.getStream(self.root_path + streamPath)
        else:
            return self.oleMessage.getStream(self.root_path + [streamPath])

    def _listDir(self):

        files = self.oleMessage.listdir()
        rootPath_ = '/'.join(self.root_path)
        inRootPath = []
        for f in files:
            f_ = '/'.join(f)
            if f_.startswith(rootPath_):
                inRootPath.append(f)
        return inRootPath

    @property
    def subject(self):
        return self._getStringStream('__substg1.0_0037')

    @property
    def header(self):
        try:
            return self._header
        except Exception:
            headerText = self._getStringStream('__substg1.0_007D')
            if headerText is not None:
                self._header = EmailParser().parsestr(headerText)
                #self._header = headerText
            else:
                self._header = None
            return self._header

    @property
    def headerStr(self):
        try:
            return self._headeStr
        except Exception:
            headerText = self._getStringStream('__substg1.0_007D')
            if headerText is not None:
                self._headeStr = headerText
                #self._header = headerText
            else:
                self._headeStr = None
            return self._headeStr

    @property
    def date(self):
        # Get the message's header and extract the date
        if self.header is None:
            return None
        else:
            return self.header['date']

    @property
    def parsedDate(self):
        return email.utils.parsedate(self.date)

    @property
    def sender(self):
        try:
            return self._sender
        except Exception:
            # Check header first
            if self.header is not None:
                headerResult = self.header["from"]
                if headerResult is not None:
                    self._sender = headerResult
                    return headerResult

            # Extract from other fields
            text = self._getStringStream('__substg1.0_0C1A')
            email = self._getStringStream('__substg1.0_0C1F')
            result = None
            if text is None:
                result = email
            else:
                result = text
                if email is not None:
                    result = result + " <" + email + ">"

            self._sender = result
            return result

    @property
    def to(self):
        try:
            return self._to
        except Exception:
            # Check header first
            if self.header is not None:
                headerResult = self.header["to"]
                if headerResult is not None:
                    self._to = headerResult
                    return headerResult

            # Extract from other fields
            # TODO: This should really extract data from the recip folders,
            # but how do you know which is to/cc/bcc?
            display = self._getStringStream('__substg1.0_0E04')
            self._to = display
            return display

    @property
    def cc(self):
        try:
            return self._cc
        except Exception:
            # Check header first
            if self.header is not None:
                headerResult = self.header["cc"]
                if headerResult is not None:
                    self._cc = headerResult
                    return headerResult

            # Extract from other fields
            # TODO: This should really extract data from the recip folders,
            # but how do you know which is to/cc/bcc?
            display = self._getStringStream('__substg1.0_0E03')
            self._cc = display
            return display

    @property
    def body(self):
        # Get the message body
        return self._getStringStream('__substg1.0_1000')

    @property
    def attachments(self):
        self._attachments = []
        attachmentDirs = []
        for dir_ in self._listDir():
            idx = len(self.root_path)
            if dir_[idx + 0].startswith('__attach') and dir_[idx + 0] not in attachmentDirs:
                if dir_[idx + 1] == '__substg1.0_37010102':
                    # Attached file
                    attachmentDirs.append(dir_[idx + 0])
                    self._attachments.append(Attachment(self, dir_[idx + 0]))
                elif dir_[idx + 1] == "__substg1.0_3701000D":
                    # Nested msg
                    attachmentDirs.append(dir_[idx + 0])
                    self._attachments.append(Message(oleMessage=self.oleMessage, root_path=dir_[idx + 0:idx + 2]))

                    #Message()
        return self._attachments

    def toJson(self):
        def xstr(s):
            return '' if s is None else str(s)

        attachmentNames = []
        # Save the attachments
      #  for attachment in self.attachments:
       #     attachmentNames.append(attachment.save())


        attachmentsJson = [a.toJson() for a in self.attachments]
        self.body = imapclient.decode_utf7(self.body)

        emailObj = {
                'from': xstr(self.sender),
                'to': xstr(self.to),
                'cc': xstr(self.cc),
                'subject': xstr(self.subject),
                'date': xstr(self.date),
                'attachments': attachmentNames,
                'body': self.body,
                'date': self.parsedDate,
                'attachments': attachmentsJson,
                'header': self.headerStr
        }
        return emailObj


if __name__ == "__main__":
    olemsg = OLEMessage(r"D:\LocalData\a189493\Desktop\Docs\code\msg\msg-extractor-master\Message suspect.msg")
    msg = Message(oleMessage=olemsg)
    #msg = Message(r"D:\LocalData\a189493\Desktop\Docs\code\msg\msg-extractor-master\Message suspect.msg")
    print msg.subject
    print msg.toJson()
    #print msg.attachments
    #msg.save()