from pathlib import Path

import openpyxl

from extractors.task_descriptions import extract_task_descriptions, attach_task_descriptions
from models.schemas import Task, OrderTypeSection


def test_extract_and_attach_task_descriptions(tmp_path: Path):
    path = tmp_path / "sample.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "6.3 OSS SO Task List"
    ws2 = wb.create_sheet("6.4 Work Order Description")
    ws2.append(["Order Type", "Task Name", "Work Order Description"])
    ws2.append(["CREATE", "ADD REWORK ACTIVITY1", "Create the work order rework activity."])
    ws2.append(["DELETE", "ADD REWORK ACTIVITY1", "Delete-flow description."])
    wb.save(path)

    descriptions = extract_task_descriptions(str(path))
    assert descriptions[("create", "add rework activity1")] == "Create the work order rework activity."

    section = OrderTypeSection(
        order_type="CREATE",
        product="MS-Edge Defender",
        title="CREATE Order Type",
        source_sheet="6.3 OSS SO Task List",
        source_rows="1-3",
        tasks=[Task(
            seq="53",
            task_name="ADD REWORK ACTIVITY1",
            it_system="CRM",
            work_group="BPM",
            source_sheet="6.3 OSS SO Task List",
            source_row=2,
        )],
    )
    matched = attach_task_descriptions([section], descriptions)
    assert matched == 1
    assert section.tasks[0].description == "Create the work order rework activity."
