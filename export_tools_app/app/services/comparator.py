"""
Service Modul: Comparator
ATURAN KETAT: MURNI Logika Backend (Tanpa dependensi Flask atau request).
"""

from typing import Dict, Any

def compare_npe_and_cipl(npe_input: Any, cipl_input: Any) -> Dict[str, Any]:
    """
    Fungsi placeholder untuk membandingkan data dari dokumen NPE/PEB dan CIPL.
    
    :param npe_input: Data/file NPE
    :param cipl_input: Data/file CIPL
    :return: Dictionary hasil perbandingan dan perincian selisih (discrepancy report)
    """
    # Dummy/Placeholder logic - Siap diisi dengan perbandingan antar DataFrame / dict
    return {
        "status": "completed",
        "npe_source": str(npe_input),
        "cipl_source": str(cipl_input),
        "match_score": "100%",
        "summary": "Data antara NPE/PEB dan CIPL Cocok (Match)",
        "comparisons": [
            {
                "field": "Jumlah Kemasan",
                "npe_value": "450 CARTON BOX",
                "cipl_value": "450 PKGS",
                "status": "MATCH",
                "badge": "success"
            },
            {
                "field": "Berat Kotor (Gross Weight)",
                "npe_value": "14,250.00 KG",
                "cipl_value": "14,250.00 KG",
                "status": "MATCH",
                "badge": "success"
            },
            {
                "field": "Berat Bersih (Net Weight)",
                "npe_value": "13,800.00 KG",
                "cipl_value": "13,800.00 KG",
                "status": "MATCH",
                "badge": "success"
            },
            {
                "field": "Consignee / Penerima",
                "npe_value": "PACIFIC TRADING CO. LTD.",
                "cipl_value": "PACIFIC TRADING CO. LTD.",
                "status": "MATCH",
                "badge": "success"
            }
        ]
    }
