from infrastructure.api_client import fetch_api_data
from domain.utils import safe_float_conversion

def fetch_data_for_designator(base_url, designator):
    response = fetch_api_data(base_url, designator)
    if not response:
        return None

    comandoAjaxSpan = response.get('comandoAjaxSpan', {})
    link_a_info = comandoAjaxSpan.get('link_a_info', {})
    link_b_info = comandoAjaxSpan.get('link_b_info', {})

    return {
        'TX A': link_a_info.get('TX', 'N/A'),
        'RX A': link_a_info.get('RX', 'N/A'),
        'Span atual A': link_a_info.get('SPAN', 'N/A'),
        'Span projetado A': link_a_info.get('SPAN_PROJ', 'N/A'),
        'Histórico A': link_a_info.get('SPAN_HIST', 'N/A'),
        'TX B': link_b_info.get('TX', 'N/A'),
        'RX B': link_b_info.get('RX', 'N/A'),
        'Span atual B': link_b_info.get('SPAN', 'N/A'),
        'Span projetado B': link_b_info.get('SPAN_PROJ', 'N/A'),
        'Histórico B': link_b_info.get('SPAN_HIST', 'N/A'),
        'KM: A <> B': safe_float_conversion(comandoAjaxSpan.get('distancia_a_b', 0))
    }
