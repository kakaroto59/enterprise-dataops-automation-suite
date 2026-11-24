"""Static AST Linter for Declarative Power Apps Canvas YAML Control Trees.

Validates automated low-code YAML generation against allowed control schemas,
Power Fx formula syntax (= prefixes, balanced string quotes and parentheses),
plain-scalar colon hazards, duplicate control identifiers, and flexbox
LayoutMinHeight clipping traps before studio import.
"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Any, Sequence

import yaml

ALLOWED_CONTROLS: frozenset[str] = frozenset(
    {
        "Classic/Button@2.2.0",
        "Classic/DropDown@2.3.1",
        "Classic/TextInput@2.3.2",
        "Gallery@2.15.0",
        "GroupContainer@1.5.0",
        "HtmlViewer@2.1.0",
        "Label@2.5.1",
        "Rectangle@2.3.0",
    }
)


def _check_formula_balance(formula: str) -> tuple[bool, int]:
    """Verify balanced double-quotes (including '' escaped "") and parentheses outside strings."""
    in_string = False
    paren_depth = 0
    idx = 0
    length = len(formula)

    while idx < length:
        ch = formula[idx]
        if ch == '"':
            if in_string and idx + 1 < length and formula[idx + 1] == '"':
                idx += 1
            else:
                in_string = not in_string
        elif not in_string:
            if ch == "(":
                paren_depth += 1
            elif ch == ")":
                paren_depth -= 1
        idx += 1

    return (not in_string), paren_depth


def _walk_control_node(
    control_name: str,
    node: Any,
    path: str,
    collected_names: list[str],
    errors: list[str],
    warnings: list[str],
) -> None:
    """Recursively validate control properties, formulas, and layout constraints."""
    if not isinstance(node, dict):
        return

    ctrl_type = node.get("Control")
    if ctrl_type not in ALLOWED_CONTROLS:
        errors.append(f"{path}/{control_name}: unverified control type {ctrl_type!r}")

    props = node.get("Properties") or {}
    if isinstance(props, dict):
        for prop_key, prop_val in props.items():
            if isinstance(prop_val, str):
                if not prop_val.startswith("="):
                    errors.append(f"{path}/{control_name}.{prop_key}: missing '=' formula prefix")
                quotes_ok, paren_bal = _check_formula_balance(prop_val)
                if not quotes_ok:
                    errors.append(f"{path}/{control_name}.{prop_key}: unbalanced string quotes")
                if paren_bal != 0:
                    errors.append(f"{path}/{control_name}.{prop_key}: unbalanced parentheses ({paren_bal})")

        # Guard against rigid LayoutMinHeight clipping inside responsive flex containers
        min_h = str(props.get("LayoutMinHeight", "")).lstrip("=").strip()
        if ctrl_type in {"GroupContainer@1.5.0", "Gallery@2.15.0"} and min_h.isdigit() and int(min_h) >= 300:
            warnings.append(
                f"{path}/{control_name}: rigid LayoutMinHeight={min_h} may clip lower cards on compact viewports"
            )

    for child in node.get("Children") or []:
        if isinstance(child, dict):
            for child_name, child_body in child.items():
                collected_names.append(child_name)
                _walk_control_node(
                    child_name,
                    child_body,
                    f"{path}/{control_name}",
                    collected_names,
                    errors,
                    warnings,
                )


def lint_yaml_tree(
    yaml_text: str,
    filename: str = "<memory>",
) -> tuple[list[str], list[str], list[str]]:
    """Lint a single declarative YAML control tree string."""
    errors: list[str] = []
    warnings: list[str] = []
    names: list[str] = []

    for line_no, raw_line in enumerate(yaml_text.splitlines(), start=1):
        stripped = raw_line.lstrip()
        if ":" in stripped:
            rhs = stripped.split(":", 1)[1]
            if rhs.lstrip().startswith("=") and ": " in rhs:
                errors.append(f"{filename}:{line_no}: plain scalar formula contains ': ' without block '|-'")

    try:
        parsed = yaml.safe_load(yaml_text)
    except yaml.YAMLError as exc:
        errors.append(f"{filename}: YAML parse error: {exc}")
        return names, errors, warnings

    for top_item in parsed or []:
        if isinstance(top_item, dict):
            for ctrl_name, ctrl_body in top_item.items():
                names.append(ctrl_name)
                _walk_control_node(ctrl_name, ctrl_body, filename, names, errors, warnings)

    return names, errors, warnings


def lint_powerapps_yaml_files(paths: Sequence[str | Path]) -> dict[str, object]:
    """Validate one or more Power Apps YAML files and flag cross-file duplicate IDs."""
    all_names: list[str] = []
    all_errors: list[str] = []
    all_warnings: list[str] = []

    for p in paths:
        file_path = Path(p)
        content = file_path.read_text(encoding="utf-8")
        names, errs, warns = lint_yaml_tree(content, filename=file_path.name)
        all_names.extend(names)
        all_errors.extend(errs)
        all_warnings.extend(warns)

    duplicates = sorted({name for name in all_names if all_names.count(name) > 1})
    if duplicates:
        all_errors.append(f"Duplicate control identifiers across tree: {duplicates}")

    return {
        "total_controls": len(all_names),
        "duplicate_controls": duplicates,
        "errors": all_errors,
        "warnings": all_warnings,
        "passed": len(all_errors) == 0,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Static AST Linter for Declarative Power Apps YAML")
    parser.add_argument("--files", nargs="+", required=True, help="One or more YAML files to validate")
    args = parser.parse_args()

    report = lint_powerapps_yaml_files(args.files)
    print(f"Controls scanned: {report['total_controls']} | Errors: {len(report['errors'])} | Warnings: {len(report['warnings'])}")
    if not report["passed"]:
        for err in report["errors"]:
            print(f"ERROR: {err}")
        raise SystemExit(1)


if __name__ == "__main__":
    main()
