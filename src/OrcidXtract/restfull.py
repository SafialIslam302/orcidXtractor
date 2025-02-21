import json
import logging
import sys

import requests
from datetime import datetime
from .constants import ORCID_PUBLIC_BASE_URL
from .utils import dictmapper, MappingRule as to
from typing import Any, Dict, List, Optional, Iterator


_logger_depth = 'INFO'

logger = logging.getLogger("#orcid#")
logger.setLevel(getattr(logging, _logger_depth))
stdout_sh = logging.StreamHandler(sys.stdout)
stdout_sh.setLevel(getattr(logging, _logger_depth))
custom_formatter = logging.Formatter(
    '%(asctime)s - %(name)s - %(levelname)s - %(message)s')
stdout_sh.setFormatter(custom_formatter)
logger.addHandler(stdout_sh)

BASE_HEADERS = {'Accept': 'application/orcid+json',
                'Content-Type': 'application/json;charset=UTF-8'}


def _parse_keywords(d: Optional[Dict[str, Any]]) -> List[str]:
    """
    Parses a dictionary to extract keywords.

    Args:
        d (Optional[Dict[str, Any]]): A dictionary containing keyword data.

    Returns:
        List[str]: A list of keywords.
    """
    if d is not None:
        return [val['content'] for val in d['keyword']]
    return []


def _parse_researcher_urls(list_l: Optional[List[Dict[str, Any]]]) -> List['Website']:
    """
    Parses a list of dictionaries to create a list of Website objects.

    Args:
        list_l (Optional[List[Dict[str, Any]]]): A list of dictionaries containing website data.

    Returns:
        List['Website']: A list of Website objects.
    """
    if list_l is not None:
        return [Website(d) for d in list_l]
    return []


def _parse_publications(list_l: Optional[List[Dict[str, Any]]]) -> List['Publication']:
    """
    Parses a list of dictionaries to create a list of Publication objects.

    Args:
        list_l (Optional[List[Dict[str, Any]]]): A list of dictionaries containing publication data.

    Returns:
        List['Publication']: A list of Publication objects.
    """
    _publications = []

    if list_l is not None:
        for d in list_l:
            path = d['work-summary'][0]['path']
            _url = '{0}{1}'.format(
                ORCID_PUBLIC_BASE_URL, path[1:])
            _res = requests.get(_url, headers=BASE_HEADERS)
            _json_body = _res.json()
            logger.debug('REQUEST (PUBLICATIONS): {0}'.format(
                json.dumps(_json_body, sort_keys=True,
                           indent=4, separators=(',', ': '))))
            _publications.append(Publication(_json_body))

    return _publications


def _parse_affiliations(list_l: Optional[List[Dict[str, Any]]]) -> List[str]:
    """
    Parses given JSON to get an affiliation
    (could be education and employment)

    Args:
        list_l (Optional[List[Dict[str, Any]]]): A list of dictionaries containing affiliation data.

    Returns:
        List[str]: A list of affiliation names.
    """

    _affiliations = []
    if list_l is not None:
        for d in list_l:
            name = d['organization']['name']
            _affiliations.append(name)
    return _affiliations


def convert_timestamp(timestamp: Optional[int]) -> str:
    """
    Convert ORCID's timestamp (milliseconds) to a
    readable date format (YYYY-MM-DD)

    Args:
        timestamp (Optional[int]): The timestamp in milliseconds.

    Returns:
        str: The formatted date string (YYYY-MM-DD), or "Unknown" if the timestamp is None.
    """
    if timestamp:
        return datetime.utcfromtimestamp(
            int(timestamp) / 1000).strftime('%Y-%m-%d')
    return "Unknown"


AuthorBase = dictmapper('AuthorBase', {
    #  'orcid'            :['orcid-profile','orcid-identifier','path'],
    'orcid': ['orcid-identifier', 'path'],
    'family_name': ['person', 'name', 'family-name', 'value'],
    'given_name': ['person', 'name', 'given-names', 'value'],
    'biography': ['person', 'biography', 'content'],
    'keywords': ['person', 'keywords'],
    'researcher_urls': ['person', 'researcher-urls', 'researcher-url'],
    'educations': ['activities-summary', 'educations', 'education-summary'],
    'employments': ['activities-summary', 'employments', 'employment-summary'],
    'last_modify_date':
        to(['history', 'last-modified-date', 'value'],
           convert_timestamp)
})

Works = dictmapper('Works', {
    'publications': to(['group'], _parse_publications),
})

PublicationBase = dictmapper('PublicationBase', {
    'title': ['title', 'title', 'value'],
    #  'url'           : ['external-ids','external-id'],
    #  'citation'      : to(['citation'], lambda l: map(CitationBase, l)
    #  if l is not None else None),
    'url': ['url'],
    'citation_value': ['citation', 'citation-value'],
    'citation_type': ['citation', 'citation-type'],
    'journal_title': ['journal-title', 'value'],
    'publicationyear': [u'publication-date', u'year', u'value'],
    'publicationtype': ['type']
})

ExternalIDBase = dictmapper('ExternalIDBase', {
    'id': ['work-external-identifier-id', 'value'],
    'type': ['work-external-identifier-type']
})

CitationBase = dictmapper('CitationBase', {
    'type': ['citation-type'],
    'value': ['citation-value']
})

WebsiteBase = dictmapper('WebsiteBase', {
    'name': ['url-name'],
    'url': ['url', 'value']
})


class Author(AuthorBase):
    """
    Represents an ORCID author.
    """
    _loaded_works = None

    def _load_works(self) -> None:
        """
        Loads the works (publications) of the author from the ORCID API.
        """
        _url = '{0}{1}/{2}'.format(ORCID_PUBLIC_BASE_URL, self.orcid, 'works')
        _res = requests.get(_url, headers=BASE_HEADERS)
        _json_body = _res.json()
        logger.debug(
            'RESPONSE (WORKS): {0}'.format(
                json.dumps(_json_body, sort_keys=True,
                           indent=4, separators=(',', ': '))))
        self._loaded_works = Works(_json_body)

    @property
    def publications(self) -> List['Publication']:
        """
        Returns the list of publications for this author.

        Returns:
            List['Publication']: A list of Publication objects.
        """
        if self._loaded_works is None:
            self._load_works()
        return self._loaded_works.publications

    @property
    def affiliations(self) -> List[Any]:
        """
        Returns the list of affiliations (educations + employments) for this author.

        Returns:
            List[Any]: A list of affiliations.
        """
        return self.educations + self.employments

    def __repr__(self) -> str:
        """
        Returns a string representation of the Author object.

        Returns:
            str: A string representation of the Author.
        """
        obj_repr = "<{} {} {}, ORCID {}>"
        return obj_repr.format(type(self).__name__,
                               self.given_name.encode('utf-8')
                               if self.given_name else 'None',
                               self.family_name.encode('utf-8')
                               if self.family_name else 'None',
                               self.orcid)

    def __str__(self) -> str:
        """
        Returns a string representation of the Author object.

        Returns:
            str: A string representation of the Author.
        """
        return self.__repr__()


class Website(WebsiteBase):
    """
    Represents a website.
    """
    def __unicode__(self) -> str:
        """
        Returns the URL of the website.

        Returns:
            str: The URL of the website.
        """
        return self.url

    def __repr__(self) -> str:
        """
        Returns a string representation of the Website object.

        Returns:
            str: A string representation of the Website.
        """
        return "<%s %s [%s]>" % (type(self).__name__, self.name, self.url)


class Citation(CitationBase):
    """
    Represents a citation.
    """
    def __unicode__(self) -> str:
        """
        Returns the text of the citation.

        Returns:
            str: The text of the citation.
        """
        return self.text

    def __repr__(self) -> str:
        """
        Returns a string representation of the Citation object.

        Returns:
            str: A string representation of the Citation.
        """
        return '<%s [type: %s]>' % (type(self).__name__, self.type)


class ExternalID(ExternalIDBase):
    """
    Represents an external identifier.
    """
    def __unicode__(self) -> str:
        """
        Returns the ID as a string.

        Returns:
            str: The ID.
        """
        return str(self.id)

    def __repr__(self) -> str:
        """
        Returns a string representation of the ExternalID object.

        Returns:
            str: A string representation of the ExternalID.
        """
        return '<%s %s:%s>' % (type(self).__name__, self.type, str(self.id))


class Publication(PublicationBase):
    """
    Represents a publication.
    """
    def __repr__(self) -> str:
        """
        Returns a string representation of the Publication object.

        Returns:
            str: A string representation of the Publication.
        """
        return '<%s "%s">' % (type(self).__name__, self.title)


#
# MAIN FUNCTIONS
#

def get(orcid_id: str) -> 'Author':
    """ Get an author based on an ORCID identifier.

    Args:
        orcid_id (str): The ORCID identifier.

    Returns:
        Author: An Author object representing the ORCID profile.
    """

    '''if sys.version_info[0] >= 3:
        unicode = str'''

    _url = '{0}{1}'.format(ORCID_PUBLIC_BASE_URL, str(orcid_id))
    _res = requests.get(_url, headers=BASE_HEADERS)

    json_body = _res.json()

    logger.debug('RESPONSE (BASE): {0}'.format(
        json.dumps(json_body, sort_keys=True, indent=4,
                   separators=(',', ': '))))

    return Author(json_body)


def search(query: str, verbose: bool = False) -> Iterator['Author']:
    """
    Searches for ORCID profiles based on a query.

    API documentation:
        https://info.orcid.org/documentation/api-tutorials/api-tutorial-searching-the-orcid-registry/

        api_example_query = {'q':'family-name:Malavolti+AND+given-names:Marco'}

    Args:
        query (str): The search query.
        verbose (bool): Enable verbose logging. Defaults to False.

    Yields:
        Author: An Author object for each search result.
    """

    if verbose:
        logger.setLevel(logging.DEBUG)
        stdout_sh.setLevel(logging.DEBUG)

    _url = '{0}{1}?q={2}'.format(ORCID_PUBLIC_BASE_URL, 'search', query)
    resp = requests.get(_url, headers=BASE_HEADERS)
    logger.debug(resp.url)
    json_body = resp.json()
    logger.debug(json_body)
    return (get(res.get('orcid-identifier', {}).get('path'))
            for res in json_body.get('result', {}))
