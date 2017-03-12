import logging
import gspread
from uuid import uuid4
from oauth2client.service_account import ServiceAccountCredentials
from pprint import pprint

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
gc = gspread.authorize(credentials)


class SheetError(Exception):
    """Raised when there are any issues with Creating or Using the Sheet class"""
    pass


class Sheet:
    s = None
    owner = None
    role = None

    def __init__(self, *args, **kwargs):
        """
        Connects to a spreadsheet. If no arguments are provided, a spreadsheet will be created with
        an arbitrary name.

        :param args: If *args is provided, it only accepts 1 value, and it used to lookup
         the spreadsheet by title
        :param kwargs: Only 1 keyword is valid at a time. Multiple keywords will result in an
         arbitrary lookup being used
            :keyword title: Looks up spreadsheet by title
            :keyword key: Looks up spreadsheet by key
            :keyword url: Looks up spreadsheet by URL
        :raises SheetError:
        """
        self.log = logging.getLogger(__name__)
        self.whoami = gc.auth._service_account_email
        if len(args) > 0:
            self._open_sheet('title', args[0])
        elif len(kwargs) > 0:
            key, value = kwargs.popitem()
            self._open_sheet(key, value)
        else:
            title = str(uuid4())
            self.log.info('No lookup value found, creating a new spreadsheet: {0}.'.format(title))
            self.s = gc.create(title)

        self._verify()
        self._initialize()

    def _open_sheet(self, key, value):
        """
        Opens the spreadsheet based on values provided
        Raises ValueError if spreadsheet can't be found.

        :param key: The lookup method to use for opening the spreadsheet
        :param value: The lookup value to find the spreadsheet by
        :return: None
        """
        try:
            if key == 'title':
                self.log.info('Opening sheet by title: {0}.'.format(value))
                self.s = gc.open(value)
            elif key == 'key':
                self.log.info('Opening sheet by key: {0}.'.format(value))
                self.s = gc.open_by_key(value)
            elif key == 'url':
                self.log.info('Opening sheet by url: {0}.'.format(value))
                self.s = gc.open_by_url(value)
            else:
                err_msg = 'Lookup key ({0}) does not match one of [title, key, url]. Aborting...'.format(key)
                self.log.error(err_msg)
                raise SheetError(err_msg)
        except gspread.exceptions.SpreadsheetNotFound:
            err_msg = 'Cannot find spreadsheet by {0}: {1}. Aborting...'.format(key, value)
            self.log.error(err_msg)
            raise SheetError(err_msg)

    def _verify(self):
        """
        Verifies at least ``edit`` access by the service account. Captures and stores the current
        file owner.

        :return: None
        :raise SheetError:
        """
        for user in self.s.list_permissions():
            if user['role'] == 'owner':
                self.owner = user['emailAddress']
            if user['emailAddress'] == self.whoami:
                self.role = user['role']

        if self.role not in ['owner', 'writer'] or self.role is None:
            err_msg = 'Insufficient privileges for service account: {0}. Aborting...'.format(self.whoami)
            self.log.error(err_msg)
            raise SheetError(err_msg)

    def _initialize(self):
        """
        Initializes the spreadsheet based on content within a ``config`` worksheet, or
        creates the ``config`` worksheet if it doesn't exist.

        The ``config`` worksheet specifies the format of all other worksheets, etc. ``Sheet._initialize`` is
        an idempotent function, and calling multiple times does not have side affects. Any worksheets no longer
        defined in ``config`` are renamed to ``_inactive-<name>``, and data is not lost.

        :return: None
        """
        if 'config' in [ws.title for ws in self.s.worksheets()]:
            # ensure indempotent matching of config
            pass
        else:
            # create config worksheet and end
            pass

    def share(self, email, role, notify=True):
        """
        Shares the spreadsheet

        :param email: Email address of the user to share with
        :param role: The role to apply to the user. ``writer` or ``reader``. See ``Sheet.change_owner`` for
            applying ownership
        :param notify: Default ``True``. Whether to notify the user or not
        :return: True if successful, False otherwise
            :rtype bool:
        """
        try:
            self.s.share(
                email,
                perm_type='user',
                role=role,
                notify=notify
            )
        except gspread.exceptions.RequestError:
            self.log.warning('Unable to share with {0}.'.format(email))
            return False
        self.log.info('Shared with {0}.'.format(email))
        return True

    def change_owner(self, new_owner_email):
        """
        Changes ownership of the spreadsheet to the supplied email address

        :param new_owner_email: Email address of the new file owner
        :return: True if successful, False otherwise
            :rtype bool:
        """
        if self.owner == self.whoami:
            try:
                self.s.share(
                    new_owner_email,
                    perm_type='user',
                    role='owner',
                    notify=False
                )
            except gspread.exceptions.RequestError:
                self.log.warning('Unable to change owner to {0}.'.format(new_owner_email))
                return False
            self.log.info('Ownership changed to {0}.'.format(new_owner_email))
            return True
        else:
            self.log.warning('Service account is not the current owner of document. Unable to change owner.')
            return False


if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    s = Sheet(url='https://docs.google.com/spreadsheets/d/17JHcvLkX0xihEztClLjs0gTyxOuaRWtR71tyAtSvnDY')
    #s = Sheet()
    #s.change_owner('wpg4665@gmail.com')

