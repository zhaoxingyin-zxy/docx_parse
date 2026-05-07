# Copyright (c) Opendatalab. All rights reserved.
import time
from io import BytesIO

from loguru import logger
from docx_parse.backend.office.model_output_to_middle_json import result_to_middle_json

from docx_parse.model.docx.main import convert_binary


def office_docx_analyze(
        file_bytes,
        image_writer=None,
        tolerant: bool = False,
):
    infer_start = time.time()

    file_stream = BytesIO(file_bytes)
    results = convert_binary(file_stream, tolerant=tolerant)

    infer_time = round(time.time() - infer_start, 2)
    safe_time = max(infer_time, 0.01)
    logger.debug(f"infer finished, cost: {infer_time}, speed: {round(len(results) / safe_time, 3)} page/s")

    middle_json = result_to_middle_json(
        results,
        image_writer,
        tolerant=tolerant,
    )
    parse_errors = getattr(results, "parse_errors", None)
    if tolerant and parse_errors:
        middle_json["_parse_errors"] = list(parse_errors)

    return middle_json, results

