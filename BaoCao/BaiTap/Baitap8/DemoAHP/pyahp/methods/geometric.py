# -*- coding: utf-8 -*-
"""pyahp.methods.geometric

This module contains the class implementing the geometric priority estimation method.
"""

import numpy as np

from pyahp.methods import Method


class GeometricMethod(Method):
    """Geometric priority estimation method
    """

    def estimate(self, preference_matrix):
        super()._check_matrix(preference_matrix)

        return np.array([np.prod(row)**(1/len(row)) for row in preference_matrix])
