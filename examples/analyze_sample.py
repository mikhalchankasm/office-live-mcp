"""Разобрать образец Excel и сохранить его «чертёж» в JSON — то же, что делает агент перед созданием аналога.

    python examples/analyze_sample.py "C:\\path\\to\\sample.xlsm" [out.blueprint.json]

Образец НЕ открывается и не меняется: сначала office_inspect_file читает его как zip (листы, макросы, правила),
затем делается копия во временной папке, на ней строится excel_describe_layout, копия закрывается и удаляется.
Excel должен быть установлен; запущенный Excel и ваши открытые книги не затрагиваются.
"""

import json
import os
import shutil
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from tests.live.mcpclient import Stdio  # noqa: E402


def main() -> int:
    if len(sys.argv) < 2:
        print(__doc__)
        return 2
    sample = os.path.abspath(sys.argv[1])
    out_path = sys.argv[2] if len(sys.argv) > 2 else os.path.splitext(sample)[0] + ".blueprint.json"
    ext = os.path.splitext(sample)[1]
    tmp = tempfile.mkdtemp(prefix="ol_blueprint_")
    c = Stdio()
    name = None
    try:
        file_info = c.call("office_inspect_file", path=sample, peek_rows=5, include_vba_source=True)
        name = c.call("excel_create_from_template", template_path=sample, new_path=os.path.join(tmp, "copy" + ext))["workbook"]
        layouts = []
        for sheet in [s["name"] for s in file_info["sheets"] if s["state"] == "visible"]:
            layouts.append(c.call("excel_describe_layout", workbook=name, sheet=sheet))
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump({"file": file_info, "layouts": layouts}, f, ensure_ascii=False, indent=1)
        print(f"Blueprint saved: {out_path}")
        for lay in layouts:
            print(f"  sheet {lay['sheet']}: {lay['data_rows']['count']} data rows, {len(lay['columns'])} columns, "
                  f"{len(lay.get('conditional_formats', []))} conditional-format rules")
        return 0
    finally:
        if name:
            c.call("excel_close_workbook", workbook=name, discard=True)
        c.close()
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    sys.exit(main())
