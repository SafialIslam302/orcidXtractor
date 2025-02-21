from typing import Optional, Any

import restfull


def get_orcid_data(orcid_id: str) -> Optional[Any]:
    """
    Retrieves ORCID data for a given ORCID ID using the restfull library.

    Args:
        orcid_id (str): The ORCID ID to retrieve data for.

    Returns:
        Optional[Any]: The ORCID data object if successful, otherwise None.
    """
    try:
        orcid_res = restfull.get(orcid_id)
        return orcid_res
    except Exception as e:
        print(f"Error retrieving data for ORCID ID {orcid_id}: {e}")
        return None
