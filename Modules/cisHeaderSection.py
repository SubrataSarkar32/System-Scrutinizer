# -*- coding: utf-8 -*-
def makeCISHeader(total_score, section):
	avgScore = round(total_score / section, 1)
	cisHeader = f'''
		<div class="blurb"></div>
		<section>
			<h1 class="reportTitle">CIS Benchmark Summary</h1>
			<div class="reportTitleDesc" style="margin-bottom: 1.5px;">
				<p class="reportTitleDescP"><b>CIS (Center for Internet Security) Benchmark Summary</b> is provides an
					overview of the device against a set of security standards and best practices established by the CIS. It
					includes an assessment of the configuration of the system and compares them to the CIS Benchmarks. The
					purpose of the report is to help organizations to understand this device security level and take action
					to improve security and compliance.</p>
			</div>
			<div class="report" id="Report">

				<table class="ReportHeader" id="Table1">
					<tbody>
						<tr>
							<th>Score:</th>
							<td><span style="font-size:120%; font-weight:bold;">{avgScore}</span> of 10</td>
						</tr>
						<tr>
							<th>Benchmark:</th>
							<td><a href="javascript:void(0)" onclick="createHelpWindow(event);"
									title="Click for complete benchmark documentation">DISA - Windows 10, Version 2.3</a>
							</td>
						</tr>
					</tbody>
				</table>
				<div title="Legend" style="float:right; margin: 0.5em 2em;">
					<img src="./assets/green.png" alt="Pass">&nbsp;=&nbsp;Pass<br>
					<img src="./assets/white.png" alt="Partial">&nbsp;=&nbsp;Partial<br>
					<img src="./assets/red.png" alt="Fail">&nbsp;=&nbsp;Fail
				</div>

				<div class="ReportToggle">
					<button class="Ctl" title="Expand" onclick="toggleAllTableBodies(this)"><img src="./assets/plus.png"
							alt=""></button>&nbsp;Expand all sections
				</div>
	'''
	return cisHeader