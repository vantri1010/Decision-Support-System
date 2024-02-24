import json
import numpy as np
from pyahp.hierarchy import AHPModel
from pyahp.methods import EigenvalueMethod
from pyahp.hierarchy import AHPCriterion

with open('risk.json') as json_model:
    model = json.load(json_model)
	
solver=EigenvalueMethod
ahp_model = AHPModel(model, EigenvalueMethod)
preference_matrices = model['preferenceMatrices']
criteria_list = model.get('criteria')
subCriteria_list = model.get('subCriteria')

criteria = [AHPCriterion(n, model, solver) for n in criteria_list]

#crit_pm = np.array(preference_matrices['criteria'])
#crit_pr = ahp_model.solver.estimate(crit_pm)

crit_attr_pr = [criterion.get_priorities() for criterion in criteria]
#if all criteria is equal (Actors = Structure = Task = Technology) then multiplier is 0.25, else multiply by "crit_pr"
attr_global_pr = [list(0.25* crit_attr_pr[i]) for i in range(len(criteria))]
print("------------")
print(subCriteria_list)
print("------------")
print(attr_global_pr)
k=input("press close to exit") 