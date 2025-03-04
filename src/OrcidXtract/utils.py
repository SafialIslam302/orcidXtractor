def dict_value_from_path(d, path):
    cur_dict = d
    for key in path[:-1]:
        cur_dict = cur_dict.get(key, {})
    return cur_dict.get(path[-1], None) if cur_dict is not None else None


def dictmapper(typename, mapping):
    """
    A factory to create `namedtuple`-like classes from a field-to-dict-path
    mapping.

    Example usage:
        Person = dictmapper('Person', {'name': ('person', 'name')})
        example_dict = {'person': {'name': 'John'}}
        john = Person(example_dict)
        assert john.name == 'John'

    If a function is specified as a mapping value instead of a dict "path",
    it will be run with the backing dict as its first argument.
    """

    def init(self, d, *args, **kwargs):
        """
        Initialize `dictmapper` classes with a dict to back getters.
        """
        self._original_dict = d

    def getter_from_dict_path(path):
        if not callable(path) and (not isinstance(path, (tuple, list)) or len(path) < 1):
            raise ValueError('Dict paths should be iterables with at least one key '
                             'or callable objects that take one argument.')

        def getter(self):
            cur_dict = self._original_dict
            if cur_dict is None:
                return None  # or raise an error depending on your needs

            if callable(path):
                return path(cur_dict)

            # Safe dictionary path access
            return dict_value_from_path(cur_dict, path)

        return getter

    def dict_value_from_path(cur_dict, path):
        """
        Helper function to safely get values from a dictionary using a path.
        """
        for key in path:
            if not cur_dict or not isinstance(cur_dict, dict):
                return None  # Return None if the path is invalid or if cur_dict is None
            cur_dict = cur_dict.get(key, None)
        return cur_dict

    # Create properties for each field in the mapping
    prop_mapping = dict((k, property(getter_from_dict_path(v)))
                        for k, v in mapping.items())
    prop_mapping['__init__'] = init

    return type(typename, tuple(), prop_mapping)


class MappingRule(object):
    def __init__(self, path, further_func=lambda x: x):
        self.path = path
        self.further_func = further_func

    def __call__(self, d):
        return self.further_func(dict_value_from_path(d, self.path))


'''def u(s):
    if sys.version_info < (3,):
        return unicode(s)
    else:
        return str(s)'''


def format_date(date_obj):
    """
    Formats a date object into a string.

    Args:
        date_obj (object): The date object to format.

    Returns:
        str: The formatted date string, or "Not Available"
        if the date object is None.
    """
    if date_obj:
        return str(date_obj)
    else:
        return "Not Available"
