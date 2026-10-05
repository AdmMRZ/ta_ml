"""Build the small, provenance-preserving clean-source workbook used by the TA.

The public module interface is ``build_clean_workbook``.  Its implementation owns
all workbook parsing, target-population filtering, validation, and presentation so
callers do not need to understand individual Excel layouts.
"""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
from math import isfinite
from pathlib import Path
from typing import Any, Iterable

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


class DataValidationError(ValueError):
    """Raised when a raw source violates the documented TA data contract."""


@dataclass(frozen=True)
class CorridorSpec:
    """Stable identity of one TA corridor across the two directly joinable sources."""

    corridor: str
    section: str
    static_sheet: str
    expected_towers: int


@dataclass(frozen=True)
class CleanBuildSummary:
    """Small result returned by the cleaning module interface."""

    output_path: Path
    tower_year_rows: int
    unique_towers: int
    static_snapshot_rows: int


RAW_LIGHTNING_FILENAME = "DATA PETIR 2021-2025 UIP3B KAL -statistik (1).xlsx"
RAW_STATIC_FILENAME = "Data Pendukung Proteksi Petir UIP3B Kal 2026.xlsx"
RAW_GEOMETRY_FILENAME = "workbook_mitigasi_r22g14_numer - SHARE.xlsx"
RAW_LIGHTNING_SHEET = "PKY-DATA PETIR"
YEARS = tuple(range(2020, 2026))

CORRIDORS = (
    CorridorSpec("PLK-KSG", "PALANGKARAYA - KASONGAN", "Kasongan-Palangkaraya", 194),
    CorridorSpec("KSG-SKS", "KASONGAN - SKS", "Kasongan-SKS", 236),
    CorridorSpec("KSG-SDN", "KASONGAN - SUDAN", "Kasongan-SudanParenggean-Sampit", 187),
)
CORRIDOR_BY_SECTION = {spec.section: spec for spec in CORRIDORS}

HISTORICAL_COLUMNS = {
    "section_raw": "Nama Section",
    "tower_name_raw": "Name",
    "coordinate_raw": "Titik Koordinat",
    "count": "Count",
    "count_negative": "Count (-)",
    "count_positive": "Count (+)",
    "positive_pct": "% Positive",
    "density": "Density",
    "min_ka": "Min kA",
    "max_ka": "Max kA",
    "mean_ka": "Mean kA",
    "exp_factor": "Exp. factor",
    "min_ka_negative": "Min kA (-)",
    "max_ka_negative": "Max kA (-)",
    "mean_ka_negative": "Mean kA (-)",
    "min_ka_positive": "Min kA (+)",
    "max_ka_positive": "Max kA (+)",
    "mean_ka_positive": "Mean kA (+)",
    "area_km2": "Area (km persegi)",
    "year": "Tahun",
    "incident_text_2021": "Data Gangguan 2021",
    "incident_text_2022": "Data Gangguan 2022",
    "incident_text_2023": "Data Gangguan 2023",
    "incident_text_2024": "Data Gangguan 2024",
    "incident_text_2025": "Data Gangguan 2025",
}

# These positions are stable across the three 2026 sheets.  Their first-row labels
# are not all unique because the source uses merged two-row headers.
STATIC_COLUMN_INDEX = {
    "grounding_leg_a_ohm": 21,
    "grounding_leg_b_ohm": 22,
    "grounding_leg_c_ohm": 23,
    "grounding_leg_d_ohm": 24,
    "grounding_total_source_ohm": 25,
    "grounding_rod_count": 26,
    "grounding_rod_length_m": 27,
    "grounding_rod_diameter_mm2": 28,
    "grounding_configuration": 29,
    "soil_resistivity_source": 30,
    "protection_angle_deg": 34,
    "circuit_configuration": 35,
    "insulator_type": 36,
    "string_l1_length": 37,
    "string_l1_bil_cfo": 38,
    "string_l2_length": 39,
    "string_l2_bil_cfo": 40,
    "string_configuration": 41,
    "tla_l1_phase": 42,
    "tla_l2_phase": 43,
    "tla_count": 44,
    "tla_type": 45,
    "tla_brand": 46,
    "tla_install_date": 47,
}

GEOMETRY_COLUMNS = {
    "elevation_m": "Elevation_m",
    "tower_height_m": "Tower_Height_m",
    "span_left_m": "Span_Left_m",
    "span_right_m": "Span_Right_m",
    "avg_neighbor_distance_m": "Avg_Neighbor_Distance_m",
    "terrain_factor": "Terrain_Factor",
    "angle_exposure_factor": "Angle_Exposure_Factor",
}


def normalize_text(value: object) -> str:
    """Return a stable key while keeping the original source value in output tables."""

    return " ".join(str(value).strip().upper().split()) if value is not None else ""


def _read_header_index(header_row: Iterable[object], required: Iterable[str]) -> dict[str, int]:
    index = {
        str(value).strip(): position
        for position, value in enumerate(header_row)
        if value is not None
    }
    missing = sorted(set(required) - set(index))
    if missing:
        raise DataValidationError(f"Missing required source columns: {', '.join(missing)}")
    return index


def _row_value(row: tuple[object, ...], index: int) -> object | None:
    return row[index] if index < len(row) else None


def _year(value: object) -> int | None:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return None
    return int(value) if float(value).is_integer() else None


def parse_indonesian_coordinates(value: object) -> tuple[float, float]:
    """Parse ``latitude;longitude`` with Indonesian decimal commas and validate bounds."""

    if not isinstance(value, str):
        raise DataValidationError(f"Coordinate is not text: {value!r}")
    parts = [part.strip().replace(" ", "").replace(",", ".") for part in value.split(";")]
    if len(parts) != 2 or not all(parts):
        raise DataValidationError(f"Expected 'latitude;longitude', got {value!r}")
    try:
        latitude, longitude = (float(part) for part in parts)
    except ValueError as error:
        raise DataValidationError(f"Invalid coordinate {value!r}") from error
    if not (-90 <= latitude <= 90 and -180 <= longitude <= 180):
        raise DataValidationError(f"Coordinate outside geographic bounds: {value!r}")
    return latitude, longitude


def _ensure_nonnegative_numeric(value: object, field_name: str, source_row: int) -> None:
    if (
        isinstance(value, bool)
        or not isinstance(value, (int, float))
        or not isfinite(value)
        or value < 0
    ):
        raise DataValidationError(
            f"{field_name} must be a non-negative number at source row {source_row}; got {value!r}"
        )


def _read_historical_rows(source_path: Path) -> list[dict[str, Any]]:
    workbook = load_workbook(source_path, read_only=True, data_only=True)
    try:
        worksheet = workbook[RAW_LIGHTNING_SHEET]
        header = next(worksheet.iter_rows(min_row=1, max_row=1, values_only=True))
        column_index = _read_header_index(header, HISTORICAL_COLUMNS.values())
        records: list[dict[str, Any]] = []

        for source_row, row in enumerate(worksheet.iter_rows(min_row=2, values_only=True), start=2):
            section_raw = _row_value(row, column_index["Nama Section"])
            section_key = normalize_text(section_raw)
            spec = CORRIDOR_BY_SECTION.get(section_key)
            year = _year(_row_value(row, column_index["Tahun"]))
            tower_name_raw = _row_value(row, column_index["Name"])
            if spec is None or year not in YEARS or not normalize_text(tower_name_raw):
                continue

            coordinate_raw = _row_value(row, column_index["Titik Koordinat"])
            latitude, longitude = parse_indonesian_coordinates(coordinate_raw)
            count = _row_value(row, column_index["Count"])
            _ensure_nonnegative_numeric(count, "Count", source_row)
            if not float(count).is_integer():
                raise DataValidationError(f"Count must be integral at source row {source_row}")

            record: dict[str, Any] = {
                "source_file": source_path.name,
                "source_sheet": RAW_LIGHTNING_SHEET,
                "source_row": source_row,
                "corridor": spec.corridor,
                "section": section_key,
                "tower_name": str(tower_name_raw).strip(),
                "tower_key": normalize_text(tower_name_raw),
                "year": year,
                "coordinate_raw": coordinate_raw,
                "latitude": latitude,
                "longitude": longitude,
            }
            for output_name, source_name in HISTORICAL_COLUMNS.items():
                if output_name not in {"section_raw", "tower_name_raw", "coordinate_raw", "year"}:
                    record[output_name] = _row_value(row, column_index[source_name])
            for field in ("count_negative", "count_positive", "density", "area_km2"):
                _ensure_nonnegative_numeric(record[field], field, source_row)
            if record["count_negative"] + record["count_positive"] != count:
                raise DataValidationError(f"Count polarity components disagree at row {source_row}")
            if record["area_km2"] <= 0:
                raise DataValidationError(f"Non-positive area at row {source_row}")
            records.append(record)
    finally:
        workbook.close()

    _validate_historical_rows(records)
    return sorted(
        records, key=lambda record: (record["corridor"], record["tower_key"], record["year"])
    )


def _validate_historical_rows(records: list[dict[str, Any]]) -> None:
    expected_rows = sum(spec.expected_towers for spec in CORRIDORS) * len(YEARS)
    if len(records) != expected_rows:
        raise DataValidationError(f"Expected {expected_rows} target rows; found {len(records)}")

    keys = [(record["tower_key"], record["year"]) for record in records]
    if len(keys) != len(set(keys)):
        raise DataValidationError("Duplicate tower_key + year found in the target population")

    for spec in CORRIDORS:
        corridor_rows = [record for record in records if record["corridor"] == spec.corridor]
        if len({record["tower_key"] for record in corridor_rows}) != spec.expected_towers:
            raise DataValidationError(
                f"{spec.corridor} does not contain {spec.expected_towers} unique towers"
            )
        year_counts = Counter(record["year"] for record in corridor_rows)
        expected_year_counts = {year: spec.expected_towers for year in YEARS}
        if dict(year_counts) != expected_year_counts:
            raise DataValidationError(
                f"{spec.corridor} does not have complete six-year coverage: {dict(year_counts)}"
            )


def _read_static_snapshot_rows(
    source_path: Path, historical_rows: list[dict[str, Any]]
) -> list[dict[str, Any]]:
    expected_keys = {
        spec.corridor: {
            row["tower_key"] for row in historical_rows if row["corridor"] == spec.corridor
        }
        for spec in CORRIDORS
    }
    workbook = load_workbook(source_path, read_only=True, data_only=True)
    records: list[dict[str, Any]] = []
    try:
        for spec in CORRIDORS:
            worksheet = workbook[spec.static_sheet]
            header = next(worksheet.iter_rows(min_row=1, max_row=1, values_only=True))
            header_index = _read_header_index(
                header, ("No.", "Nama Section", "Name", "Titik Koordinat")
            )
            corridor_records: list[dict[str, Any]] = []

            for source_row, row in enumerate(
                worksheet.iter_rows(min_row=3, values_only=True), start=3
            ):
                source_number = _row_value(row, header_index["No."])
                section_key = normalize_text(_row_value(row, header_index["Nama Section"]))
                tower_name_raw = _row_value(row, header_index["Name"])
                if not isinstance(source_number, (int, float)) or section_key != spec.section:
                    continue

                coordinate_raw = _row_value(row, header_index["Titik Koordinat"])
                latitude, longitude = parse_indonesian_coordinates(coordinate_raw)
                record: dict[str, Any] = {
                    "source_file": source_path.name,
                    "source_sheet": spec.static_sheet,
                    "source_row": source_row,
                    "source_file_label_year": 2026,
                    "corridor": spec.corridor,
                    "section": section_key,
                    "tower_name": str(tower_name_raw).strip(),
                    "tower_key": normalize_text(tower_name_raw),
                    "coordinate_2026_raw": coordinate_raw,
                    "latitude_2026": latitude,
                    "longitude_2026": longitude,
                }
                for output_name, position in STATIC_COLUMN_INDEX.items():
                    record[output_name] = _row_value(row, position)
                corridor_records.append(record)

            _validate_static_snapshot_rows(spec, corridor_records, expected_keys[spec.corridor])
            records.extend(corridor_records)
    finally:
        workbook.close()

    return sorted(records, key=lambda record: (record["corridor"], record["tower_key"]))


def _geometry_number(value: object, field: str, source_row: int) -> float:
    if isinstance(value, bool):
        raise DataValidationError(f"Invalid {field} at INPUT_TOWER row {source_row}: {value!r}")
    try:
        number = float(value)
    except (TypeError, ValueError) as error:
        raise DataValidationError(
            f"Invalid {field} at INPUT_TOWER row {source_row}: {value!r}"
        ) from error
    if not isfinite(number) or number < 0:
        raise DataValidationError(f"Invalid {field} at INPUT_TOWER row {source_row}: {value!r}")
    return number


def _join_geometry(source_path: Path, static_rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    workbook = load_workbook(source_path, read_only=True, data_only=True)
    try:
        worksheet = workbook["INPUT_TOWER"]
        header = next(worksheet.iter_rows(max_row=1, values_only=True))
        columns = _read_header_index(header, ("Section", "Tower_Name", *GEOMETRY_COLUMNS.values()))
        geometry: dict[tuple[str, str], dict[str, Any]] = {}
        for source_row, row in enumerate(worksheet.iter_rows(min_row=2, values_only=True), start=2):
            corridor = normalize_text(_row_value(row, columns["Section"]))
            tower_key = normalize_text(_row_value(row, columns["Tower_Name"]))
            key = (corridor, tower_key)
            if corridor not in {spec.corridor for spec in CORRIDORS}:
                continue
            if not tower_key or key in geometry:
                raise DataValidationError(
                    f"Missing/duplicate geometry identity at INPUT_TOWER row {source_row}"
                )
            record = {
                "geometry_source_file": source_path.name,
                "geometry_source_sheet": "INPUT_TOWER",
                "geometry_source_row": source_row,
            }
            for field, source_field in GEOMETRY_COLUMNS.items():
                record[field] = _geometry_number(
                    _row_value(row, columns[source_field]), field, source_row
                )
            if record["tower_height_m"] <= 0 or record["avg_neighbor_distance_m"] <= 0:
                raise DataValidationError(
                    f"Non-positive height/distance at INPUT_TOWER row {source_row}"
                )
            geometry[key] = record
    finally:
        workbook.close()

    expected = {(r["corridor"], r["tower_key"]) for r in static_rows}
    if set(geometry) != expected:
        raise DataValidationError(
            f"Geometry registry mismatch: missing={sorted(expected - set(geometry))[:3]}, "
            f"unexpected={sorted(set(geometry) - expected)[:3]}"
        )
    registry = []
    for static_row in static_rows:
        row = dict(static_row)
        row["static_source_file"] = row.pop("source_file")
        row["static_source_sheet"] = row.pop("source_sheet")
        row["static_source_row"] = row.pop("source_row")
        row.update(geometry[(row["corridor"], row["tower_key"])])
        registry.append(row)
    return registry


def _attach_grounding_reference(source_path: Path, registry_rows: list[dict[str, Any]]) -> None:
    """Compare the snapshot with Sheet4 without replacing either source's values."""

    expected_names = {row["tower_key"] for row in registry_rows}
    workbook = load_workbook(source_path, read_only=True, data_only=True)
    references: dict[str, tuple[int, tuple[object, ...]]] = {}
    try:
        worksheet = workbook["Sheet4"]
        header = next(worksheet.iter_rows(max_row=1, values_only=True))
        columns = _read_header_index(
            header,
            (
                "Name",
                "Grounding Leg A",
                "Grounding Leg B",
                "Grounding Leg C",
                "Grounding Leg D",
                "Grounding Total",
            ),
        )
        for source_row, row in enumerate(worksheet.iter_rows(min_row=2, values_only=True), 2):
            tower_key = normalize_text(_row_value(row, columns["Name"]))
            if tower_key not in expected_names:
                continue
            if tower_key in references:
                raise DataValidationError(f"Duplicate target tower in Sheet4: {tower_key}")
            references[tower_key] = (source_row, row)
    finally:
        workbook.close()

    if set(references) != expected_names:
        raise DataValidationError(
            f"Sheet4 grounding reference lacks {len(expected_names - set(references))} tower names"
        )
    for registry in registry_rows:
        source_row, source = references[registry["tower_key"]]
        registry["grounding_sheet4_source_file"] = source_path.name
        registry["grounding_sheet4_source_sheet"] = "Sheet4"
        registry["grounding_sheet4_source_row"] = source_row
        registry["grounding_abc_matches_sheet4"] = all(
            registry[f"grounding_leg_{leg}_ohm"] == source[columns[f"Grounding Leg {leg.upper()}"]]
            for leg in "abc"
        )
        registry["grounding_leg_d_sheet4_reference_ohm"] = source[columns["Grounding Leg D"]]
        registry["grounding_total_sheet4_reference_ohm"] = source[columns["Grounding Total"]]
        registry["grounding_d_matches_sheet4"] = (
            registry["grounding_leg_d_ohm"] == registry["grounding_leg_d_sheet4_reference_ohm"]
        )
        registry["grounding_total_matches_sheet4"] = (
            registry["grounding_total_source_ohm"]
            == registry["grounding_total_sheet4_reference_ohm"]
        )


def _audit_2020_coordinates(
    historical_rows: list[dict[str, Any]], registry_rows: list[dict[str, Any]]
) -> list[dict[str, Any]]:
    """Record exact coordinate mappings as candidates, not verified Count corrections."""

    by_coordinate: dict[tuple[float, float], dict[str, Any]] = {}
    for registry in registry_rows:
        coordinate = (registry["latitude_2026"], registry["longitude_2026"])
        if coordinate in by_coordinate:
            raise DataValidationError(f"Registry has duplicate coordinates: {coordinate}")
        by_coordinate[coordinate] = registry

    audit: list[dict[str, Any]] = []
    matched_keys: set[tuple[str, str]] = set()
    for row in historical_rows:
        if row["year"] != 2020:
            continue
        matched = by_coordinate.get((row["latitude"], row["longitude"]))
        if matched is not None:
            matched_key = (matched["corridor"], matched["tower_key"])
            if matched_key in matched_keys:
                raise DataValidationError(
                    f"Two 2020 rows map to the same coordinate: {matched_key}"
                )
            matched_keys.add(matched_key)
        audit.append(
            {
                "source_file": row["source_file"],
                "source_sheet": row["source_sheet"],
                "source_row": row["source_row"],
                "source_corridor": row["corridor"],
                "source_tower_key": row["tower_key"],
                "source_count_2020": row["count"],
                "latitude_2020": row["latitude"],
                "longitude_2020": row["longitude"],
                "match_status": "exact_coordinate" if matched is not None else "unmatched",
                "matched_corridor": matched["corridor"] if matched is not None else None,
                "matched_tower_key": matched["tower_key"] if matched is not None else None,
                "matched_registry_source_row": matched["static_source_row"]
                if matched is not None
                else None,
                "count_usage_rule": (
                    "used_after_user_directed_exact_coordinate_reassignment"
                    if matched is not None
                    else "unassigned_no_exact_coordinate"
                ),
            }
        )
    return audit


def _reconcile_historical_rows(
    historical_rows: list[dict[str, Any]],
    registry_rows: list[dict[str, Any]],
    coordinate_audit: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """Re-key exact 2020 coordinate matches; leave unresolved Counts unassigned."""

    registry = {(row["corridor"], row["tower_key"]): row for row in registry_rows}
    audit_by_source_row = {row["source_row"]: row for row in coordinate_audit}
    if len(audit_by_source_row) != 617:
        raise DataValidationError("2020 coordinate audit lacks unique source rows")

    reconciled: list[dict[str, Any]] = []
    matched_2020: set[tuple[str, str]] = set()
    for source in historical_rows:
        row = dict(source)
        row["source_corridor"] = source["corridor"]
        row["source_section"] = source["section"]
        row["source_tower_name"] = source["tower_name"]
        row["source_tower_key"] = source["tower_key"]
        row["coordinate_origin"] = "source"
        if source["year"] == 2020:
            audit = audit_by_source_row[source["source_row"]]
            if audit["match_status"] != "exact_coordinate":
                continue  # The complete raw record is preserved in coordinate_2020_audit.
            key = (audit["matched_corridor"], audit["matched_tower_key"])
            if key in matched_2020 or key not in registry:
                raise DataValidationError(f"Invalid/duplicate 2020 coordinate target: {key}")
            matched_2020.add(key)
            target = registry[key]
            row["corridor"] = target["corridor"]
            row["section"] = target["section"]
            row["tower_name"] = target["tower_name"]
            row["tower_key"] = target["tower_key"]
            row["reconciliation_status"] = "exact_coordinate_reassigned"
        else:
            row["reconciliation_status"] = "source_name_coordinate_verified"
        reconciled.append(row)

    for key in sorted(set(registry) - matched_2020):
        target = registry[key]
        placeholder = {field: None for field in historical_rows[0]}
        placeholder.update(
            {
                "corridor": target["corridor"],
                "section": target["section"],
                "tower_name": target["tower_name"],
                "tower_key": target["tower_key"],
                "year": 2020,
                "latitude": target["latitude_2026"],
                "longitude": target["longitude_2026"],
                "source_corridor": None,
                "source_section": None,
                "source_tower_name": None,
                "source_tower_key": None,
                "coordinate_origin": "verified_registry_placeholder",
                "reconciliation_status": "unresolved_2020_no_count",
            }
        )
        reconciled.append(placeholder)

    expected = {(corridor, tower_key, year) for corridor, tower_key in registry for year in YEARS}
    actual = [(row["corridor"], row["tower_key"], row["year"]) for row in reconciled]
    if len(actual) != 3702 or set(actual) != expected or len(actual) != len(set(actual)):
        raise DataValidationError("Reconciled tower-year population is incomplete or duplicated")
    for row in reconciled:
        if row["source_row"] is None:
            row["coordinate_registry_match"] = None
            continue
        target = registry[(row["corridor"], row["tower_key"])]
        row["coordinate_registry_match"] = (
            row["latitude"] == target["latitude_2026"]
            and row["longitude"] == target["longitude_2026"]
        )
        if not row["coordinate_registry_match"]:
            raise DataValidationError(
                f"Source coordinate disagrees after reconciliation: {row['tower_key']}, {row['year']}"
            )
    return sorted(reconciled, key=lambda row: (row["corridor"], row["tower_key"], row["year"]))


def _validate_static_snapshot_rows(
    spec: CorridorSpec, records: list[dict[str, Any]], expected_keys: set[str]
) -> None:
    keys = [record["tower_key"] for record in records]
    if len(keys) != spec.expected_towers:
        raise DataValidationError(
            f"{spec.static_sheet} should contain {spec.expected_towers} selected source rows; found {len(keys)}"
        )
    if len(keys) != len(set(keys)):
        raise DataValidationError(f"Duplicate tower key in 2026 snapshot: {spec.static_sheet}")
    if set(keys) != expected_keys:
        missing = sorted(expected_keys - set(keys))
        unexpected = sorted(set(keys) - expected_keys)
        raise DataValidationError(
            f"2026 snapshot identity mismatch for {spec.corridor}; "
            f"missing={missing[:3]}, unexpected={unexpected[:3]}"
        )


def _source_notes() -> list[dict[str, Any]]:
    return [
        {
            "source": RAW_LIGHTNING_FILENAME,
            "role_in_clean_workbook": "Historical tower-year lightning observations",
            "included_table": "tower_year; coordinate_2020_audit; tower_registry Sheet4 reference",
            "time_scope": "2020-2025",
            "join_or_usage_rule": "2020 records are re-keyed by exact unique coordinate for 609 towers on user instruction; 8 Counts stay unassigned in coordinate_2020_audit, with missing 2020 observations in tower_year. Raw Excel is unchanged.",
        },
        {
            "source": RAW_STATIC_FILENAME,
            "role_in_clean_workbook": "Tower protection/grounding candidate attributes",
            "included_table": "tower_registry",
            "time_scope": "File labeled 2026; grounding measurement dates unknown",
            "join_or_usage_rule": "Exact normalized-name match to 617 towers; no measurement date per grounding value. Leg D/Total disagree with Sheet4 for some towers and remain raw source values.",
        },
        {
            "source": "Elevasi Tower Palangka-Kasongan-SKS.xlsx",
            "role_in_clean_workbook": "Potential elevation, height, and span reference",
            "included_table": "Not joined",
            "time_scope": "Undated reference",
            "join_or_usage_rule": "Tower identities are sequence/T labels, not verified tower-name keys; raw file remains authoritative.",
        },
        {
            "source": "Data Elevasi Kasongan-SKS.xlsx",
            "role_in_clean_workbook": "Potential KSG-SKS elevation/grounding reference",
            "included_table": "Not joined",
            "time_scope": "Undated reference",
            "join_or_usage_rule": "x/y and tower identity formatting need provenance and coordinate validation before any join.",
        },
        {
            "source": "DENSITY & PEAK CURRENT 2021 - 2025 (PLK-KSG-SKS-Kurun).xlsx",
            "role_in_clean_workbook": "Aggregate corridor lightning reference",
            "included_table": "Not joined",
            "time_scope": "2021-2025",
            "join_or_usage_rule": "Aggregate data are not tower-level; corridor definitions must not be assumed equivalent.",
        },
        {
            "source": "workbook_mitigasi_r22g14_numer - SHARE.xlsx",
            "role_in_clean_workbook": "INPUT_TOWER geometry source and separate PLN comparator",
            "included_table": "tower_registry: INPUT_TOWER geometry only",
            "time_scope": "Engineering workbook snapshot",
            "join_or_usage_rule": "INPUT_TOWER geometry joined by corridor + tower name; derived PLN outputs excluded from ML.",
        },
    ]


def _write_table(workbook: Workbook, title: str, rows: list[dict[str, Any]]) -> None:
    worksheet = workbook.create_sheet(title)
    if not rows:
        worksheet.append(["No rows"])
        return

    headers = list(rows[0])
    worksheet.append(headers)
    for row in rows:
        worksheet.append([row.get(header) for header in headers])

    header_fill = PatternFill("solid", fgColor="1F4E78")
    for cell in worksheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions
    worksheet.sheet_view.showGridLines = False
    for column_index, header in enumerate(headers, start=1):
        sample_lengths = [
            len(str(row.get(header, ""))) for row in rows[:200] if row.get(header) is not None
        ]
        sample_width = max([len(str(header)), *sample_lengths])
        worksheet.column_dimensions[get_column_letter(column_index)].width = min(
            max(sample_width + 2, 12), 34
        )


def _write_readme(workbook: Workbook, summary: CleanBuildSummary) -> None:
    worksheet = workbook.create_sheet("README")
    worksheet.sheet_view.showGridLines = False
    rows = [
        ("TA clean-source workbook", "v2"),
        ("Purpose", "Clean, traceable source layer. It is not yet the final ML feature matrix."),
        (
            "Historical table",
            f"tower_year: {summary.tower_year_rows} rows; one row per tower-year (2020-2025).",
        ),
        (
            "Registry",
            f"tower_registry: {summary.static_snapshot_rows} towers; valid PLN geometry plus 2026 protection snapshot.",
        ),
        ("Population", f"{summary.unique_towers} towers: PLK-KSG 194, KSG-SKS 236, KSG-SDN 187."),
        (
            "Key",
            "tower_key is a normalized name for joining; preserve tower_name for display and traceability.",
        ),
        (
            "Important",
            "The 2026 protection snapshot is not merged into historical tower_year rows; geometry provenance is in tower_registry.",
        ),
        (
            "Important",
            "All 617 original 2020 names disagree with registry coordinates. 609 observations are re-keyed by exact unique GPS; 8 Count values remain unassigned and visible in coordinate_2020_audit.",
        ),
        (
            "Coordinate audit",
            "coordinate_2020_audit preserves original 2020 names, Counts, coordinates, and source rows. Reconciled tower_year uses only exact one-to-one GPS matches.",
        ),
        (
            "Grounding audit",
            "tower_registry compares 2026 Leg D/Total values with Sheet4; discrepancies are flagged, not silently corrected.",
        ),
        (
            "Important",
            "PKY grounding fields are excluded from tower_year because their source formulas need separate provenance validation.",
        ),
        (
            "Important",
            "Density, Count (+/-), positive percentage, same-year current statistics, and PLN outputs are not legal features for predicting Count in that same year.",
        ),
        ("Rebuild", "python -m ta_ml.cli --overwrite"),
    ]
    for row in rows:
        worksheet.append(row)
    worksheet["A1"].font = Font(bold=True, color="FFFFFF")
    worksheet["B1"].font = Font(bold=True, color="FFFFFF")
    for cell in worksheet[1]:
        cell.fill = PatternFill("solid", fgColor="1F4E78")
    worksheet.column_dimensions["A"].width = 24
    worksheet.column_dimensions["B"].width = 120
    for row in worksheet.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)


def build_clean_workbook(
    *, source_dir: Path, output_path: Path, overwrite: bool = False
) -> CleanBuildSummary:
    """Create one validated clean workbook without modifying any raw Excel source.

    The output contains a tower-year historical table and a separate 2026 static
    snapshot table.  The separation is an invariant: the builder never performs an
    implicit historical join from the 2026 snapshot.
    """

    source_dir = source_dir.resolve()
    output_path = output_path.resolve()
    lightning_path = source_dir / RAW_LIGHTNING_FILENAME
    static_path = source_dir / RAW_STATIC_FILENAME
    geometry_path = source_dir / RAW_GEOMETRY_FILENAME
    missing_sources = [
        str(path) for path in (lightning_path, static_path, geometry_path) if not path.is_file()
    ]
    if missing_sources:
        raise DataValidationError(f"Missing raw source workbook(s): {', '.join(missing_sources)}")
    if output_path in {lightning_path.resolve(), static_path.resolve(), geometry_path.resolve()}:
        raise DataValidationError("Output path cannot replace a raw source workbook")
    if output_path.exists() and not overwrite:
        raise FileExistsError(
            f"Output already exists: {output_path}. Re-run with --overwrite to replace it."
        )

    historical_rows = _read_historical_rows(lightning_path)
    static_rows = _read_static_snapshot_rows(static_path, historical_rows)
    registry_rows = _join_geometry(geometry_path, static_rows)
    _attach_grounding_reference(lightning_path, registry_rows)
    coordinate_audit = _audit_2020_coordinates(historical_rows, registry_rows)
    historical_rows = _reconcile_historical_rows(historical_rows, registry_rows, coordinate_audit)
    summary = CleanBuildSummary(
        output_path=output_path,
        tower_year_rows=len(historical_rows),
        unique_towers=len({row["tower_key"] for row in historical_rows}),
        static_snapshot_rows=len(static_rows),
    )

    workbook = Workbook()
    workbook.remove(workbook.active)
    _write_readme(workbook, summary)
    _write_table(workbook, "tower_year", historical_rows)
    _write_table(workbook, "tower_registry", registry_rows)
    _write_table(workbook, "coordinate_2020_audit", coordinate_audit)
    _write_table(workbook, "source_notes", _source_notes())

    output_path.parent.mkdir(parents=True, exist_ok=True)
    temporary_path = output_path.with_suffix(".tmp.xlsx")
    try:
        workbook.save(temporary_path)
        temporary_path.replace(output_path)
    finally:
        workbook.close()
        if temporary_path.exists():
            temporary_path.unlink()
    return summary
