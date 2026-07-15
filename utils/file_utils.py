"""
utils/file_utils.py
-------------------
Helpers for building output paths and managing downloaded report files.
"""

import os
from datetime import datetime


def resolve_path(template: str, now: datetime | None = None) -> str:
    """
    Replace date/time placeholders in *template* with values from *now*.

    Supported placeholders
    ----------------------
    ``{YYYY}``  – 4-digit year
    ``{MM}``    – 2-digit month
    ``{DD}``    – 2-digit day
    ``{HH}``    – 2-digit hour (24 h)
    ``{MIN}``   – 2-digit minute
    ``{SS}``    – 2-digit second

    Parameters
    ----------
    template:
        Path or filename template string, e.g.
        ``"output/{YYYY}/{MM}/VL06F_{YYYY}{MM}{DD}.xlsx"``.
    now:
        Point in time to use for substitution.  Defaults to
        :func:`datetime.now` if *None*.

    Returns
    -------
    str
        Resolved path string.
    """
    if now is None:
        now = datetime.now()

    return (
        template
        .replace("{YYYY}", now.strftime("%Y"))
        .replace("{MM}",   now.strftime("%m"))
        .replace("{DD}",   now.strftime("%d"))
        .replace("{HH}",   now.strftime("%H"))
        .replace("{MIN}",  now.strftime("%M"))
        .replace("{SS}",   now.strftime("%S"))
    )


def ensure_dir(path: str) -> str:
    """
    Create *path* (and any missing parents) if it does not exist.

    Parameters
    ----------
    path:
        Directory path to create.

    Returns
    -------
    str
        The same *path* that was passed in (convenient for chaining).
    """
    os.makedirs(path, exist_ok=True)
    return path


def build_output_path(output_dir_template: str,
                      file_name_template: str,
                      now: datetime | None = None) -> str:
    """
    Combine *output_dir_template* and *file_name_template* into a full file
    path, resolve date placeholders, and ensure the directory exists.

    Parameters
    ----------
    output_dir_template:
        Directory template, e.g. ``"output/{YYYY}/{MM}"``.
    file_name_template:
        File name template, e.g. ``"VL06F_{YYYY}{MM}{DD}_{HH}{MIN}{SS}.xlsx"``.
    now:
        Timestamp used for placeholder substitution.

    Returns
    -------
    str
        Absolute file path ready for writing.
    """
    if now is None:
        now = datetime.now()

    resolved_dir  = resolve_path(output_dir_template, now)
    resolved_file = resolve_path(file_name_template,  now)
    ensure_dir(resolved_dir)
    return os.path.join(resolved_dir, resolved_file)
