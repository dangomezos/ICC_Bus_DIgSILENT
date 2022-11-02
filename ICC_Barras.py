# -*- coding: utf-8 -*-
"""
Autor: Daniel Gómez

Estudios Eléctricos S.A.S

"""

import powerfactory as pf

##### Inicia la aplicación #####
app=pf.GetApplication()
app.ClearOutputWindow()
app.EchoOff()

##### Obtener los datos dentro del ComPython ######
script=app.GetCurrentScript()
path = script.Path
file_name = script.Filename

oShcs = script.GetContents('*.ComShc') ### Toma todos los elementos de Shc
oScenarios = script.GetContents('Scenarios')[0].GetAll('IntScenario') ## Toma todos los escenario
oBuses = script.GetContents('Buses')[0] ### Toma todas las barras donde se hacen las fallas
app.PrintPlain(oShcs)

def run_shc(oShc, scenario):
	oShc.iopt_allbus = 0
	oShc.shcobj=oBuses
	app.PrintPlain('Escenario: {} -> tipo:{} -> con Rf:{} en {}##### '.format(scenario, oShc.iopt_shc, oShc.Rf, oBuses))

	oShc.Execute()

	for bus in script.GetContents('Buses')[0].GetAll('ElmTerm'):
		if '3psc' == oShc.iopt_shc:
			f.write("%s\t%s\t%.1f\t%s\t%2.1f\t%0.3f\t%s\t%s\t%s\t%s\n"%(scenario.loc_name, (bus.loc_name),bus.uknom , oShc.iopt_shc, oShc.Rf, 
			bus.GetAttribute('m:Ikss'),'N/A' ,'N/A', 'N/A', 'N/A'))
		elif 'spgf' == oShc.iopt_shc or '2pgf' == oShc.iopt_shc or '2psc' == oShc.iopt_shc:
			f.write("%s\t%s\t%.1f\t%s\t%2.1f\t%s\t%0.3f\t%0.3f\t%0.3f\t%0.3f\n"%(scenario.loc_name, (bus.loc_name), bus.uknom,
			oShc.iopt_shc, oShc.Rf, 'N/A',bus.GetAttribute('m:Ikss:A'), bus.GetAttribute('m:Ikss:B'), bus.GetAttribute('m:Ikss:C'), bus.GetAttribute('m:I0x3')))
	 
# bus.cpSubstat.loc_name+"_"+bus.loc_name

def main():
	for oScenario in oScenarios:
		valid_scenario=oScenario.Activate()
		# assert valid_scenario == 0, f"Invalid Scenario" 
		for oSch in oShcs:
			run_shc(oSch, oScenario)

if __name__ == '__main__': ### EntryPoint ###
	with open(path+file_name+'.txt', 'w', encoding='utf-8') as f:
		f.write("%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\n"%('Esenario', 'Elemento', 'Tensión [kV]', 
                                                      'Tipo_falla', 'Resistencia', 'Ikss', 'IkssA', 'IkssB','IkssC','I0x3'))
		main()