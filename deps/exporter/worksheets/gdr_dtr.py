"""
#
# gdr_dtr.py - exporter worksheets - STDF Viewer
#
# Author: falbany (fork)
# Created Date: November 12th 2026
# -----
# Last Modified: Thu Nov 12 2026
# Modified By: falbany (fork)
# -----
# Copyright (c) 2026 falbany (fork)
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.
#

GDR & DTR Summary worksheet writer (moved verbatim from generator.py).
"""

from deps.uic_stdExporter import ReportSelection
from deps.exporter.sheets import plain_table_spec
from deps.exporter.theme import TAB_COLOR_HEXES
from deps.exporter.writer import render_plain_table


class GdrDtrMixin:
    """GDR & DTR Summary worksheet writer."""

    def writeGDR_DTR(self):
        GDR_DTR_Sheet = self.sheetDict.get(ReportSelection.GDR_DTR, None)
        if GDR_DTR_Sheet is None:
            return

        dataTable = self.waitForTableData(isFileInfo=False)
        if self.forceQuit:
            return
        # remove \n that added besides values
        dataTable = [[ele.strip("\n") for ele in dataRow] for dataRow in dataTable]

        spec = plain_table_spec(
            self.tr("GDR & DTR Summary"),
            [self.tr("Record Type"), self.tr("Value"), self.tr("Approx. Location")],
            dataTable,
            self.tr,
        )
        spec["tab_color"] = TAB_COLOR_HEXES["gdr"]
        spec["repeat_header_row"] = 1
        render_plain_table(GDR_DTR_Sheet, spec, self._theme)
        self._autofit_columns(GDR_DTR_Sheet)

        self.loopCount += 1
        self.sendProgress()
