"""Domain models and metadata mappings for the ASK & Z1F Vendor Package Monitoring Dashboard."""

from dataclasses import dataclass, field, asdict
from typing import Dict, List, Optional, Any


# Standard 12 Key Milestones across WPR and PSR
STANDARD_MILESTONES = [
    (1, "Contract Award", "po_issued"),
    (2, "Kick Off Meeting", "kom"),
    (3, "VDRL Submission", "vd_submission"),
    (4, "Approval of Key Documents", "approval_key_docs"),
    (5, "Delivery of Major Equipment/Material", "delivery_materials"),
    (6, "Start of Production", "start_production"),
    (7, "FAT - Hardware", "fat_hardware"),
    (8, "FAT - Software", "fat_software"),
    (9, "Punch Clearance - Hardware", "punch_clearance"),
    (10, "Ready for Shipment", "ready_shipment"),
    (11, "Receipt at Worksite", "receipt_worksite"),
    (12, "Approval of Vendor Final Documentation", "final_documentation"),
]


# Canonical ASK Package Catalog (maps VDRL / KOM / VDLR codes to PSR RFQ / Package Names)
ASK_PACKAGE_METADATA: Dict[str, Dict[str, str]] = {
    "PLC": {"name": "PLC System", "discipline": "INS", "rfq": "RFQ-INS-037", "psr_name": "PLC System", "tag": "WP-01 / WP-02"},
    "ABV": {"name": "Actuated Ball Valve", "discipline": "INS", "rfq": "RFQ-INS-024", "psr_name": "Actuated Ball Valve", "tag": "WP-01 / WP-02"},
    "ACBV": {"name": "Actuated Compact Ball Valve", "discipline": "INS", "rfq": "RFQ-INS-025", "psr_name": "Actuated Compact Ball Valve", "tag": "WP-01 / WP-02"},
    "ACV": {"name": "Actuated Choke Valve", "discipline": "INS", "rfq": "RFQ-INS-023", "psr_name": "Actuated Choke Valve", "tag": "WP-01 / WP-02"},
    "CV": {"name": "Control Valves", "discipline": "INS", "rfq": "RFQ-INS-021", "psr_name": "Control Valve", "tag": "WP-01 / WP-02"},
    "SRV": {"name": "Self Regulator Valve", "discipline": "INS", "rfq": "RFQ-INS-022", "psr_name": "Self Regulator Valve", "tag": "WP-01 / WP-02"},
    "WHCP": {"name": "Wellhead Control Panel (WHCP)", "discipline": "INS", "rfq": "RFQ-INS-028", "psr_name": "Wellhead Control Panel (WHCP)", "tag": "WHCP-01 / WHCP-02"},
    "PDC": {"name": "Offshore Pedestal Crane", "discipline": "MEC", "rfq": "RFQ-MEC-001", "psr_name": "Offshore Pedestal Crane", "tag": "CR-07001 (WP01 & WP02)"},
    "CIP": {"name": "Chemical Injection Pumps Package", "discipline": "MEC", "rfq": "RFQ-MEC-006", "psr_name": "Chemical Injection Pumps Package", "tag": "SK-09001 (WP01 & WP02)"},
    "CDP": {"name": "Closed Drain Sump Pump & Recovery Water Pump Package", "discipline": "MEC", "rfq": "RFQ-MEC-014", "psr_name": "Closed Drain Sump Pump And Recovery Water Pump Package", "tag": "WP-01 / WP-02"},
    "DTP": {"name": "Diesel Transfer Pump Package", "discipline": "MEC", "rfq": "RFQ-MEC-015", "psr_name": "Diesel Transfer Pump Package", "tag": "WP-01 / WP-02"},
    "OPDP": {"name": "Open Drain Caisson Pump Package", "discipline": "MEC", "rfq": "RFQ-MEC-016", "psr_name": "Open Drain Caisson Pump Package", "tag": "WP-01 / WP-02"},
    "PV": {"name": "Pressure Vessels Package", "discipline": "PV", "rfq": "RFQ-MEC-004", "psr_name": "Test And Production Separator", "tag": "V-01001 / V-10001 / V-10002"},
    "TSP": {"name": "Test Separator", "discipline": "PV", "rfq": "RFQ-MEC-004", "psr_name": "Test And Production Separator", "tag": "V-01001 (WP01 & WP02)"},
    "PSP": {"name": "Production Separator", "discipline": "PV", "rfq": "RFQ-MEC-004", "psr_name": "Test And Production Separator", "tag": "V-10001 (WP02)"},
    "SKT": {"name": "Skim Tank Vessel", "discipline": "PV", "rfq": "RFQ-MEC-010", "psr_name": "Skim Tank Vessel", "tag": "V-10002 (WP02)"},
    "CDS": {"name": "Closed Drain Sump Vessel", "discipline": "PV", "rfq": "RFQ-MEC-011", "psr_name": "Closed Drain Sump Vessel", "tag": "V-08001 (WP01 & WP02)"},
    "DGS": {"name": "Degasser Drum", "discipline": "PV", "rfq": "RFQ-MEC-009", "psr_name": "Degasser Drum", "tag": "WP-02"},
    "DOL": {"name": "Deoiler Hydro Cyclone Package", "discipline": "PV", "rfq": "RFQ-MEC-008", "psr_name": "Deoiler Hydro Cyclone Package", "tag": "WP-02"},
    "IGF": {"name": "Induced Gas Floatation (IGF) Package", "discipline": "PV", "rfq": "RFQ-MEC-008", "psr_name": "Deoiler Hydro Cyclone Package", "tag": "WP-02"},
    "IGFIL": {"name": "Instrument Gas Filter Coalescer", "discipline": "PV", "rfq": "RFQ-MEC-003", "psr_name": "Instrument Gas Filter Coalescer", "tag": "WP-01 / WP-02"},
    "TANK": {"name": "Atmospheric Storage Tanks", "discipline": "PV", "rfq": "RFQ-MEC-013", "psr_name": "Storage Tank", "tag": "WP-01 / WP-02"},
    "SOLR": {"name": "Solar Power System / ACDC / Ex-Charger / NavAid", "discipline": "ELE", "rfq": "RFQ-ELE-018", "psr_name": "Solar Power System", "tag": "WP-01 / WP-02"},
    "BV": {"name": "Manual Ball Valves", "discipline": "PIP", "rfq": "RFQ-PIP-048", "psr_name": "Ball Valve (Flange End)", "tag": "WP-01 / WP-02"},
    "GLV": {"name": "Manual Globe & Check Valves", "discipline": "PIP", "rfq": "RFQ-PIP-050", "psr_name": "Globe Valve (Flange End)", "tag": "WP-01 / WP-02"},
    "PLS": {"name": "Position Limit Switches", "discipline": "INS", "rfq": "RFQ-INS-040", "psr_name": "Position Limit Switches", "tag": "WP-01 / WP-02"},
    "PSV": {"name": "Safety Relief Valves (PSV)", "discipline": "INS", "rfq": "RFQ-INS-026", "psr_name": "Safety Relief Valves", "tag": "WP-01 / WP-02"},
    "PG": {"name": "Pressure & Temperature Gauges (PG/TG)", "discipline": "INS", "rfq": "RFQ-INS-029", "psr_name": "Pressure Gauge", "tag": "WP-01 / WP-02"},
    "PLT": {"name": "Primary Steel - Plates (Plate)", "discipline": "STR", "rfq": "RFQ-STR-055", "psr_name": "Primary Steel - Plates", "tag": "WP-01 / WP-02"},
    "BEAM": {"name": "Primary Steel - Beams", "discipline": "STR", "rfq": "RFQ-STR-056", "psr_name": "Primary Steel - Beam", "tag": "WP-01 / WP-02"},
    "SMT": {"name": "Structural Seamless Tubular", "discipline": "STR", "rfq": "RFQ-STR-057", "psr_name": "Primary Seamless Structure Pipe", "tag": "WP-01 / WP-02"},
}


# Canonical Z1F Package Catalog (maps WPR & PSR packages for Z1F)
Z1F_PACKAGE_METADATA: Dict[str, Dict[str, str]] = {
    "CCV": {"name": "Compact Check Valve", "discipline": "PIP", "rfq": "RFQ-PIP-004", "psr_name": "Check valve (Compact)", "tag": "ZWP20/21/22/23-VCC-J54"},
    "PDC": {"name": "Offshore Pedestal Crane", "discipline": "MEC", "rfq": "RFQ-MA-001", "psr_name": "Offshore Pedestal Crane", "tag": "ZWP20 / ZWP22 / ZWP23"},
    "CDS": {"name": "Closed Drain Sump Vessel", "discipline": "PV", "rfq": "RFQ-MA-002-1", "psr_name": "Closed Drain Sump Vessel", "tag": "ZWP20 / ZWP21 / ZWP22 / ZWP23"},
    "TSP": {"name": "Test Separator", "discipline": "PV", "rfq": "RFQ-MA-002-2", "psr_name": "Test Separator", "tag": "ZWP20 / ZWP21 / ZWP22 / ZWP23"},
    "IGFIL": {"name": "Instrument Gas Filter Coalescer", "discipline": "PV", "rfq": "RFQ-MA-002-3", "psr_name": "Instrument Gas Filter Coalescer", "tag": "ZWP20 / ZWP21 / ZWP22 / ZWP23"},
    "CDP": {"name": "Closed Drain Sump Pump Skids", "discipline": "MEC", "rfq": "RFQ-MA-003", "psr_name": "Closed Drain Sump Pump skids", "tag": "ZWP20 / ZWP21 / ZWP22 / ZWP23"},
    "CIP": {"name": "Chemical Injection Pump Skids", "discipline": "MEC", "rfq": "RFQ-MA-004", "psr_name": "Chemical injection pump skids", "tag": "ZWP20 / ZWP21 / ZWP22 / ZWP23"},
    "TANK": {"name": "Chemical Storage Tank", "discipline": "PV", "rfq": "RFQ-MA-005", "psr_name": "Chemical storage tank", "tag": "ZWP20 / ZWP21 / ZWP22 / ZWP23"},
    "SOLR": {"name": "Solar Power System", "discipline": "ELE", "rfq": "RFQ-EL-001", "psr_name": "Solar Power System", "tag": "ZWP20 / ZWP21 / ZWP22 / ZWP23"},
    "ACV": {"name": "Actuated Choke Valve", "discipline": "INS", "rfq": "RFQ-IN-001", "psr_name": "Actuated Choke Valve", "tag": "ZWP20 / ZWP21 / ZWP22 / ZWP23"},
    "ABV": {"name": "Actuated Ball Valve", "discipline": "INS", "rfq": "RFQ-IN-002", "psr_name": "Actuated ball valve", "tag": "ZWP20 / ZWP21 / ZWP22 / ZWP23"},
    "ACBV": {"name": "Actuated Compact Ball Valve", "discipline": "INS", "rfq": "RFQ-IN-003", "psr_name": "Actuated Compact Ball Valve", "tag": "ZWP20 / ZWP21 / ZWP22 / ZWP23"},
    "PSV": {"name": "Safety Relief Valves", "discipline": "INS", "rfq": "RFQ-IN-004", "psr_name": "Safety Relief Valves", "tag": "ZWP20 / ZWP21 / ZWP22 / ZWP23"},
    "WHCP": {"name": "Wellhead Control Panel (WHCP)", "discipline": "INS", "rfq": "RFQ-IN-006", "psr_name": "Wellhead Control Panel", "tag": "ZWP20 / ZWP21 / ZWP22 / ZWP23"},
    "PLC": {"name": "PLC System", "discipline": "INS", "rfq": "RFQ-IN-007", "psr_name": "PLC system", "tag": "ZWP20 / ZWP21 / ZWP22 / ZWP23"},
    "CV": {"name": "Control Valve", "discipline": "INS", "rfq": "RFQ-IN-008-1", "psr_name": "Control valve", "tag": "ZWP20 / ZWP21 / ZWP22 / ZWP23"},
    "SRV": {"name": "Self-Actuated Pressure Control Valves", "discipline": "INS", "rfq": "RFQ-IN-008-2", "psr_name": "Self-Actuated Pressure Control Valves", "tag": "ZWP20 / ZWP21 / ZWP22 / ZWP23"},
}


@dataclass
class ManufacturerDetail:
    """Element 3: Manufacturer Detail strictly populated from KOM/ files (blank if no KOM)."""
    available: bool = False
    manufacturer_name: str = ""
    vendor_code: str = ""
    factory_location: str = ""
    kom_date: str = ""
    po_number: str = ""
    po_date: str = ""
    equipment_scope: str = ""
    key_contacts: str = ""
    kom_source_file: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

