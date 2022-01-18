# Constraint-Based-Performance-Scheduler

Given a set of events, a set of months over which to distribute them, and a set of constraints and priorities to satisfy, this program schedules the performances optimally
while meeting the constraints set in place.

compiler.py is responsible for cleaning and consolidating several databases into a single database that the scheduler (assigner.py) then uses in order to output a final set
of schedulings that is optimal for the given inputs. The code is built primarily on Python, with Pandas used for the data processing modules.
