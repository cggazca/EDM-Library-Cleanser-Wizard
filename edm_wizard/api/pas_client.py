"""
Part Aggregation Service (PAS) API Client

Implements SearchAndAssign matching algorithm from legacy Java tool
"""

import re
import time
from datetime import datetime, timedelta
import requests

from ..utils.constants import (
    PAS_API_URL,
    PAS_AUTH_URL,
    PAS_SEARCH_PROVIDER_ID,
    PAS_SEARCH_PROVIDER_VERSION,
    PAS_PROPERTY_MANUFACTURER_NAME,
    PAS_PROPERTY_MANUFACTURER_PN,
    PAS_PROPERTY_DATASHEET_URL,
    PAS_PROPERTY_FINDCHIPS_URL,
    PAS_PROPERTY_LIFECYCLE_STATUS,
    PAS_PROPERTY_LIFECYCLE_STATUS_CODE,
    PAS_PROPERTY_PART_ID,
    MATCH_TYPE_FOUND,
    MATCH_TYPE_MULTIPLE,
    MATCH_TYPE_NEED_REVIEW,
    MATCH_TYPE_NONE,
    MATCH_TYPE_ERROR,
    DEFAULT_MAX_MATCHES
)


class PASAPIClient:
    """Part Aggregation Service API Client with OAuth 2.0 authentication"""

    def __init__(self, client_id, client_secret, max_matches=DEFAULT_MAX_MATCHES):
        """
        Initialize PAS API client with credentials

        Args:
            client_id: OAuth client ID
            client_secret: OAuth client secret
            max_matches: Maximum matches to return per part (default: 10)
        """
        self.client_id = client_id
        self.client_secret = client_secret
        self.pas_url = PAS_API_URL
        self.auth_url = PAS_AUTH_URL
        self.access_token = None
        self.token_expires_at = None
        self.max_matches = max_matches

    def _get_access_token(self):
        """
        Get or refresh the OAuth access token

        Returns:
            Access token string

        Raises:
            requests.HTTPError: If authentication fails
        """
        if self.access_token and self.token_expires_at:
            if datetime.now() < self.token_expires_at:
                return self.access_token

        # Request new token
        auth = (self.client_id, self.client_secret)
        auth_data = {
            'grant_type': 'client_credentials',
            'scope': 'sws.icarus.api.read'
        }
        headers = {
            'Content-Type': 'application/x-www-form-urlencoded'
        }

        response = requests.post(
            self.auth_url,
            auth=auth,
            data=auth_data,
            headers=headers,
            timeout=10
        )
        response.raise_for_status()

        token_data = response.json()
        self.access_token = token_data['access_token']
        expires_in = token_data.get('expires_in', 7200)
        # Refresh 1 minute before expiry
        self.token_expires_at = datetime.now() + timedelta(seconds=expires_in - 60)

        return self.access_token

    def search_part(self, manufacturer_pn, manufacturer):
        """
        Search for a part using PAS API with SearchAndAssign matching algorithm

        Implements the exact matching logic from SearchAndAssignApp.java:
        1. Exact match: PartNumber AND ManufacturerName both exact
        2. Partial manufacturer match: PartNumber exact, ManufacturerName contains
        3. Alphanumeric-only match: Strip special chars, compare alphanumeric only
        4. Leading zero suppression: Remove leading zeros and compare
        5. PartNumber-only search: If manufacturer empty/unknown

        Args:
            manufacturer_pn: Part number to search for
            manufacturer: Manufacturer name (optional)

        Returns:
            Tuple of (result_dict, match_type)
            - result_dict: {'matches': [...], 'raw_matches': [...]} or {'error': str}
            - match_type: One of MATCH_TYPE_* constants
        """
        try:
            # Perform PAS search
            search_results = self._perform_pas_search(manufacturer_pn, manufacturer)

            if 'error' in search_results:
                return {'error': search_results['error']}, MATCH_TYPE_ERROR

            parts = search_results.get('results', [])

            if not parts:
                return {'matches': []}, MATCH_TYPE_NONE

            # Apply SearchAndAssign matching algorithm
            match_result = self._apply_searchandassign_matching(
                manufacturer_pn, manufacturer, parts
            )

            return match_result

        except Exception as e:
            return {'error': str(e)}, MATCH_TYPE_ERROR

    def _perform_pas_search(self, manufacturer_pn, manufacturer):
        """
        Perform PAS API parametric search

        Uses parametric/search endpoint with property-based filters for more
        accurate results than free-text search.

        Args:
            manufacturer_pn: Part number to search for
            manufacturer: Manufacturer name (optional)

        Returns:
            Dict with 'results' list or 'error' string
        """
        try:
            token = self._get_access_token()

            # Parametric search endpoint
            endpoint = f'/api/v2/search-providers/{PAS_SEARCH_PROVIDER_ID}/{PAS_SEARCH_PROVIDER_VERSION}/parametric/search'

            headers = {
                'Authorization': f'Bearer {token}',
                'Content-Type': 'application/json',
                'X-Siemens-Correlation-Id': f'corr-{int(time.time() * 1000)}',
                'X-Siemens-Session-Id': f'session-{int(time.time())}',
                'X-Siemens-Ebs-User-Country-Code': 'US',
                'X-Siemens-Ebs-User-Currency': 'USD'
            }

            # Build filter based on whether manufacturer is provided
            if manufacturer and manufacturer.strip():
                # Two-parameter search: AND filter for both Part Number and Manufacturer
                search_filter = {
                    "__logicalOperator__": "And",
                    "__expression__": "LogicalExpression",
                    "left": {
                        "__valueOperator__": "SmartMatch",
                        "__expression__": "ValueExpression",
                        "propertyId": PAS_PROPERTY_MANUFACTURER_NAME,
                        "term": manufacturer
                    },
                    "right": {
                        "__valueOperator__": "SmartMatch",
                        "__expression__": "ValueExpression",
                        "propertyId": PAS_PROPERTY_MANUFACTURER_PN,
                        "term": manufacturer_pn
                    }
                }
                page_size = 10  # Two-parameter search
            else:
                # One-parameter search: filter by Part Number only
                search_filter = {
                    "__valueOperator__": "SmartMatch",
                    "__expression__": "ValueExpression",
                    "propertyId": PAS_PROPERTY_MANUFACTURER_PN,
                    "term": manufacturer_pn
                }
                page_size = 50  # One-parameter search

            request_body = {
                "searchParameters": {
                    "partClassId": "76f2225d",  # Root part class
                    "customParameters": {},
                    "outputs": [
                        PAS_PROPERTY_MANUFACTURER_NAME,
                        PAS_PROPERTY_MANUFACTURER_PN,
                        PAS_PROPERTY_DATASHEET_URL,
                        PAS_PROPERTY_FINDCHIPS_URL,
                        PAS_PROPERTY_LIFECYCLE_STATUS,
                        PAS_PROPERTY_LIFECYCLE_STATUS_CODE,
                        PAS_PROPERTY_PART_ID
                    ],
                    "sort": [],
                    "paging": {
                        "requestedPageSize": page_size
                    },
                    "filter": search_filter
                }
            }

            # Collect all results (handle pagination)
            all_results = []
            url = f"{self.pas_url}{endpoint}"

            while True:
                response = requests.post(
                    url,
                    headers=headers,
                    json=request_body,
                    timeout=60
                )

                if response.status_code == 401:
                    # Token expired, retry once
                    self.access_token = None
                    self.token_expires_at = None
                    token = self._get_access_token()
                    headers['Authorization'] = f'Bearer {token}'
                    response = requests.post(
                        url,
                        headers=headers,
                        json=request_body,
                        timeout=60
                    )

                response.raise_for_status()
                result = response.json()

                if not result.get('success', False):
                    error = result.get('error', {})
                    error_msg = error.get('message', 'Unknown error')
                    return {'error': error_msg}

                # Add results from this page
                if result.get('result') and result['result'].get('results'):
                    all_results.extend(result['result']['results'])

                # Check for next page
                next_page_token = result.get('result', {}).get('nextPageToken')
                if not next_page_token:
                    break

                # Prepare next page request
                endpoint = f'/api/v2/search-providers/{PAS_SEARCH_PROVIDER_ID}/{PAS_SEARCH_PROVIDER_VERSION}/parametric/get-next-page'
                url = f"{self.pas_url}{endpoint}"
                request_body = {
                    "pageToken": next_page_token
                }

            return {
                'results': all_results,
                'totalCount': len(all_results)
            }

        except Exception as e:
            return {'error': str(e)}

    def _apply_searchandassign_matching(self, edm_pn, edm_mfg, parts):
        """
        Apply the SearchAndAssign matching algorithm from Java code

        Mirrors Java implementation:
        - If manufacturer is NOT empty/Unknown: Try exact + partial + alphanumeric + zero suppression
        - If no match OR manufacturer IS empty/Unknown: Search by PartNumber only

        Args:
            edm_pn: EDM part number
            edm_mfg: EDM manufacturer name
            parts: List of part data from PAS API

        Returns:
            Tuple of (result_dict, match_type)
        """
        pattern = re.compile(r'[^A-Za-z0-9]')
        matches = []
        result_record = None

        # ========== STEP 1: Search with Manufacturer (if provided) ==========
        if edm_mfg and edm_mfg not in ['', 'Unknown']:
            # 1a. Exact match on BOTH PartNumber AND ManufacturerName
            for part_data in parts:
                part = part_data.get('searchProviderPart', {})
                pas_pn = part.get('manufacturerPartNumber', '')
                pas_mfg = part.get('manufacturerName', '')
                if pas_pn == edm_pn and pas_mfg == edm_mfg:
                    matches.append(part_data)

            if len(matches) > 1:
                return self._format_match_result(matches, MATCH_TYPE_MULTIPLE)
            elif len(matches) == 1:
                return self._format_match_result(matches, MATCH_TYPE_FOUND)

            # 1b. Partial match on ManufacturerName
            matches.clear()
            for part_data in parts:
                part = part_data.get('searchProviderPart', {})
                pas_pn = part.get('manufacturerPartNumber', '')
                pas_mfg = part.get('manufacturerName', '')
                if pas_pn == edm_pn and edm_mfg in pas_mfg:
                    matches.append(part_data)

            if len(matches) > 1:
                return self._format_match_result(matches, MATCH_TYPE_MULTIPLE)
            elif len(matches) == 1:
                return self._format_match_result(matches, MATCH_TYPE_FOUND)

            # 1c. Alphanumeric-only match
            matches.clear()
            edm_pn_alpha = pattern.sub('', edm_pn)
            for part_data in parts:
                part = part_data.get('searchProviderPart', {})
                pas_pn = part.get('manufacturerPartNumber', '')
                pas_mfg = part.get('manufacturerName', '')
                pas_pn_alpha = pattern.sub('', pas_pn)
                # Check both part number AND manufacturer (partial match OK)
                if pas_pn_alpha == edm_pn_alpha and (pas_mfg == edm_mfg or edm_mfg in pas_mfg):
                    matches.append(part_data)

            if len(matches) == 0:
                # 1d. Leading zero suppression
                edm_pn_no_zeros = edm_pn_alpha.lstrip('0')
                for part_data in parts:
                    part = part_data.get('searchProviderPart', {})
                    pas_pn = part.get('manufacturerPartNumber', '')
                    pas_mfg = part.get('manufacturerName', '')
                    pas_pn_alpha = pattern.sub('', pas_pn)
                    pas_pn_no_zeros = pas_pn_alpha.lstrip('0')
                    # Check both part number AND manufacturer (partial match OK)
                    if pas_pn_no_zeros == edm_pn_no_zeros and (pas_mfg == edm_mfg or edm_mfg in pas_mfg):
                        matches.append(part_data)

                if len(matches) == 1:
                    return self._format_match_result(matches, MATCH_TYPE_FOUND)
                elif len(matches) > 1:
                    return self._format_match_result([matches[0]], MATCH_TYPE_FOUND)
            else:
                if len(matches) >= 1:
                    return self._format_match_result([matches[0]], MATCH_TYPE_FOUND)

        # ========== STEP 2: Search by PartNumber only ==========
        # Triggered if: manufacturer is empty/Unknown OR no matches found in Step 1
        if result_record is None and (not edm_mfg or edm_mfg in ['', 'Unknown'] or len(matches) == 0):
            matches.clear()
            all_results = list(parts)

            if len(parts) == 0:
                return {'matches': []}, MATCH_TYPE_NONE

            # Special case: If exactly 1 result from PAS search
            if len(parts) == 1:
                return self._format_match_result(parts, MATCH_TYPE_NEED_REVIEW)

            # Multiple results from PAS - try to narrow down by PartNumber
            # 2a. Exact PartNumber match
            for part_data in parts:
                part = part_data.get('searchProviderPart', {})
                pas_pn = part.get('manufacturerPartNumber', '')
                if pas_pn == edm_pn:
                    matches.append(part_data)

            if len(matches) == 0:
                # 2b. Alphanumeric-only match
                edm_pn_alpha = pattern.sub('', edm_pn)
                for part_data in parts:
                    part = part_data.get('searchProviderPart', {})
                    pas_pn = part.get('manufacturerPartNumber', '')
                    pas_pn_alpha = pattern.sub('', pas_pn)
                    if pas_pn_alpha == edm_pn_alpha:
                        matches.append(part_data)

                if len(matches) == 0:
                    # 2c. Leading zero suppression
                    edm_pn_no_zeros = edm_pn_alpha.lstrip('0')
                    for part_data in parts:
                        part = part_data.get('searchProviderPart', {})
                        pas_pn = part.get('manufacturerPartNumber', '')
                        pas_pn_alpha = pattern.sub('', pas_pn)
                        pas_pn_no_zeros = pas_pn_alpha.lstrip('0')
                        if pas_pn_no_zeros == edm_pn_no_zeros:
                            matches.append(part_data)

                    if len(matches) == 0:
                        # No matches - return all as Multiple
                        return self._format_match_result(all_results, MATCH_TYPE_MULTIPLE)
                    elif len(matches) == 1:
                        return self._format_match_result(matches, MATCH_TYPE_NEED_REVIEW)
                    else:
                        # Multiple matches - take first
                        return self._format_match_result([matches[0]], MATCH_TYPE_FOUND)
                else:
                    if len(matches) == 1:
                        return self._format_match_result(matches, MATCH_TYPE_NEED_REVIEW)
                    elif len(matches) > 1:
                        return self._format_match_result(matches, MATCH_TYPE_MULTIPLE)
            else:
                if len(matches) == 1:
                    return self._format_match_result(matches, MATCH_TYPE_NEED_REVIEW)
                elif len(matches) > 1:
                    # Multiple exact matches
                    return self._format_match_result(matches, MATCH_TYPE_MULTIPLE)

        # No matches found (fallback)
        return {'matches': []}, MATCH_TYPE_NONE

    def _format_match_result(self, part_data_list, match_type):
        """
        Format the match result in a consistent way

        Args:
            part_data_list: List of part data from PAS API
            match_type: One of MATCH_TYPE_* constants

        Returns:
            Tuple of (result_dict, match_type)
        """
        matches = []
        for part_data in part_data_list:
            part = part_data.get('searchProviderPart', {})
            mpn = part.get('manufacturerPartNumber', '')
            mfg = part.get('manufacturerName', '')

            # Extract DataProviderID (partId) - External Content ID
            part_id = part.get('partId', '')

            # Extract lifecycle status and Findchips URL from properties
            properties = part.get('properties', {}).get('succeeded', {})
            lifecycle_status = properties.get(PAS_PROPERTY_LIFECYCLE_STATUS, '')
            lifecycle_code = properties.get(PAS_PROPERTY_LIFECYCLE_STATUS_CODE, '')
            findchips_url = properties.get(PAS_PROPERTY_FINDCHIPS_URL, '')

            # If findchips_url is a URL object, extract the value
            if isinstance(findchips_url, dict) and '__complex__' in findchips_url:
                findchips_url = findchips_url.get('value', '')

            # Create match entry with metadata
            match_entry = {
                'mpn': mpn,
                'mfg': mfg,
                'lifecycle_status': lifecycle_status,
                'lifecycle_code': lifecycle_code,
                'external_id': part_id,
                'findchips_url': findchips_url,
                'match_string': f"{mpn}@{mfg}"
            }
            matches.append(match_entry)

        # Limit matches to user-configured maximum
        return {
            'matches': matches[:self.max_matches],
            'raw_matches': part_data_list[:self.max_matches]
        }, match_type
