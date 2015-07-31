from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.types import DateTime
from sqlalchemy.types import String, Unicode
from sqlalchemy.types import Integer
from sqlalchemy.types import Boolean
from sqlalchemy import Column, ForeignKey, create_engine
from sqlalchemy.orm import relationship, backref, sessionmaker
from ExtractMsg import Message as MessageParser
from ExtractMsg import Attachment as AttachmentParser
import logging
from datetime import datetime
import re
import json

logging.basicConfig(level=logging.DEBUG)
Base = declarative_base()

class Message(Base):
    __tablename__ = 'message'
    id = Column(Integer, primary_key=True)
    sender = Column(String)
    sender_email = Column(String)
    to = Column(String)
    cc = Column(String)
    subject = Column(Unicode)
    header = Column(String)
    urls = Column(String)
    date = Column(DateTime)
    body = Column(String)
    spf_pass = Column(Boolean)
    distinct_senders_in_header = Column(Integer)
    from_mismatch_header = Column(Boolean)

    internal_mail = Column(Boolean)
    attachments = relationship("Attachment", backref="message")

    parent_id = Column(Integer, ForeignKey('message.id'))
    nested_messages = relationship("Message", backref=backref('parent', remote_side=[id]))

    def __init__(self, msgFilePath=None, msgParser=None):
        if msgFilePath is not None:
            self.msg_parser = MessageParser(msgFilePath=msgFilePath)
        elif msgParser is not None:
            self.msg_parser = msgParser
        else:
            raise Exception("No path or msgParser given")
        self.sender = self.msg_parser.sender
        try:
            self.sender_email = re.search(r"([a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+)", self.sender).group() #Parse according to RFC5322
        except AttributeError:
            self.sender_email = None
        self.to = self.msg_parser.to
        self.cc = self.msg_parser.cc
        self.subject = self.msg_parser.subject
        self.header = self.msg_parser.headerStr
        self.body = self.msg_parser.body
        self.urls = json.dumps(self.extract_urls(self.body))
        date = self.msg_parser.parsedDate
        self.date = datetime(
            year=date[0],
            month=date[1],
            day=date[2],
            hour=date[3],
            minute=date[4],
            second=date[5]
        )

        for attachment in self.msg_parser.attachments:
            if isinstance(attachment, MessageParser):
                self.nested_messages.append(Message(msgParser=attachment))
            elif isinstance(attachment, AttachmentParser):
                self.attachments.append(Attachment(attachment))

        self.score_mail()
        pass

    @staticmethod
    def extract_urls(html):
        if not html:
            return
        urls = re.findall(r'(http[s]?://.*?)(?:>| |(?:\r\n){2})', html, flags=re.DOTALL|re.MULTILINE)
        urls = [url.replace('\r\n','') for url in urls]
        for url in urls:

        return urls

    def score_mail(self):
        status, spf_senders = self.spf()
        envelope_from, x_sender = self.check_sender()
        if not self.internal_mail:
            all_senders = spf_senders + envelope_from + x_sender
            all_senders = list(set(all_senders))
            if len(all_senders) > 1:
                logging.warning("/!\ Multiple senders declared : %s"%all_senders)
                self.distinct_senders_in_header = len(all_senders)
            if self.sender_email not in all_senders:
                logging.warning("/!\ sender email %s is not present in server generated headers (%s), email must be forged"%(self.sender_email, all_senders))
                self.from_mismatch_header = True
        else:
            #Internal
            pass

    def check_sender(self):
        envelope_from = re.findall("envelope-from=\"(.*?)\"", self.header, flags=re.MULTILINE|re.DOTALL)
        envelope_from = list(set(envelope_from))
        logging.info("enveloppe from : %s" % envelope_from)
        x_sender = re.findall("x-sender=\"(.*?)\"", self.header, flags=re.MULTILINE|re.DOTALL)
        x_sender = list(set(x_sender))
        logging.info("xsender : %s" % x_sender)
        return envelope_from, x_sender

    def spf(self):
        spfs = re.findall("Received-SPF: (Pass|None) (\(.*?\.MYDOMAIN\.fr: .*?\))", self.header, flags=re.MULTILINE|re.DOTALL)
        senders = []
        spf_pass = False
        for spf in spfs:
            logging.debug("SPF : %s, %s"%spf)
            if spf[0] == "Pass":
                spf_pass = True
                logging.info("SPF is OK : %s, %s"%spf)
            sender = re.findall("\S*?@[^) ]*",spf[1])
            if sender is not None:
                senders += sender
        if not spf_pass:
            if not "Received-SPF" in self.header:
                logging.info("Internal email : %s"%self.header)
                self.internal_mail = True
                spf_pass = True
            else:
                logging.info("SPF is KO : %s"%spfs)
        senders = list(set(senders))
        logging.info("SPF Senders : %s"%senders)
        self.spf_pass = spf_pass
        return spf_pass, senders

class Attachment(Base):
    __tablename__ = 'attachments'
    id = Column(Integer, primary_key=True)
    short_name = Column(String)
    long_name = Column(String)
    magic = Column(String)
    sha1 = Column(String)
    message_id = Column(Integer, ForeignKey('message.id'))
    risky = Column(Boolean)
    risky_ext = {'.bat': '', '.bin': '',  '.cmd': '',  '.com': '',  '.cpl': '',  '.dll': '',  '.doc': '',  '.docb': '',  '.docm': '',  '.docx': '',  '.dot': '',  '.dotm': '',  '.dotx': '',  '.exe': '',  '.hta': '',  '.htm': '',  '.html': '',
        '.jar': '',  '.msc': '',  '.msi': '',  '.msp': '',  '.mst': '',  '.pdf': '',  '.pif': '',  '.pot': '',  '.potm': '',  '.potx': '',  '.ppam': '',  '.pps': '',  '.ppsm': '',  '.ppsx': '',  '.ppt': '',  '.pptm': '',  '.pptx': '',
        '.ps1': '',  '.ps1xml': '',  '.ps2': '',  '.ps2xml': '',  '.psc1': '',  '.psc2': '',  '.reg': '',  '.rgs': '',  '.scr': '',  '.sct': '',  '.shb': '',  '.shs': '',  '.sldm': '',  '.sldx': '',  '.vb': '',  '.vba': '',  '.vbe': '',
        '.vbs': '',  '.vbscript': '',  '.ws': '',  '.wsh': '',  '.xla': '',  '.xlam': '',  '.xll': '',  '.xlm': '',  '.xls': '',  '.xlsb': '',  '.xlsm': '',  '.xlsx': '',  '.xlt': '',  '.xltm': '',  '.xltx': '',  '.xlw': '',  '.zip': ''
    }

    def __init__(self, attachment_obj):
        self.short_name = attachment_obj.shortFilename
        self.long_name = attachment_obj.longFilename
        self.sha1 = attachment_obj.sha1
        data = attachment_obj.data
        self.is_risky()

    def is_risky(self):
        ext = "." + self.long_name.split(".")[-1]
        if ext in self.risky_ext:
            self.risky = True

engine = create_engine('sqlite:///db.sqlite', echo=True)
Base.metadata.create_all(engine)
Session = sessionmaker(bind=engine)
session=Session()
import glob
for mail in glob.glob(u"mails/*.msg"):
    msg = Message(msgFilePath=mail)
    session.add(msg)
session.commit()
pass