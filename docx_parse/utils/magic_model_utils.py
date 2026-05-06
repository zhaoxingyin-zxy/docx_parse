from typing import Callable


def tie_up_category_by_index(
    get_subjects_func: Callable,
    get_objects_func: Callable,
    extract_subject_func: Callable | None = None,
    extract_object_func: Callable | None = None,
    object_block_type: str = "object",
    include_bbox: bool = True,
):
    """Associate nearby subject/object blocks by original block index.

    The standalone DOCX path calls this with include_bbox=False because Office
    blocks do not carry page-space bounding boxes.
    """

    subjects = get_subjects_func()
    objects = get_objects_func()
    extract_subject_func = extract_subject_func or (lambda x: x)
    extract_object_func = extract_object_func or (lambda x: x)

    result = {
        i: {
            "sub_bbox": extract_subject_func(subject),
            "obj_bboxes": [],
            "sub_idx": i,
        }
        for i, subject in enumerate(subjects)
    }
    if not subjects:
        return []

    object_indices = {obj["index"] for obj in objects}

    def effective_index_diff(obj_index: int, sub_index: int) -> int:
        start, end = min(obj_index, sub_index), max(obj_index, sub_index)
        return (end - start) - sum(1 for idx in range(start + 1, end) if idx in object_indices)

    for obj in objects:
        obj_index = obj["index"]
        best_subject_idx = min(
            range(len(subjects)),
            key=lambda i: effective_index_diff(obj_index, subjects[i]["index"]),
        )
        result[best_subject_idx]["obj_bboxes"].append(extract_object_func(obj))

    ret = list(result.values())
    ret.sort(key=lambda x: x["sub_idx"])
    return ret
