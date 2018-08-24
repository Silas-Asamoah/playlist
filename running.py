# -*- coding: utf-8 -*-
"""
Created on Tue Jul 31 12:28:21 2018

@author: Peg
"""

import holoviews as hv
renderer = hv.renderer('matplotlib').instance(fig='svg', holomap='gif')
renderer.save(my_object, 'example_I', style=dict(Image={'cmap':'jet'}))