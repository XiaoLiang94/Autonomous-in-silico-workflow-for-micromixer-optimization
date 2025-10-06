# Autonomous-in-silico-workflow-for-micromixer-optimization

**Overview**

This project proposes an autonomous in-silico workflow to accelerate the optimization of high-performance obstacle-based micromixers.

The workflow integrated a multi-objective genetic algorithm for guiding the optimization interation, SolidWorks for 3D modeling micromixers, STAR-CCM+ for performing CFD simulations, and Excel for mixing performance evaluation and data storage, enabling a fully automated and human-free optimization process. The schematic diagram is shown below.


<p align="center">
  <img src="images/Figure.jpg" width="200" alt="Micromixer Design">
</p>

The repository contains all the main code and files to run a demonstration example, which employs 3 generations and 2 solutions per generation.

**Folder Structure**

  **Main program:**

main.py — the script to run the workflow.

  **Macro files:**

Creating3D.bas — used to automatically generate 3D micromixer models with defined obstacles in SolidWorks, based on the Excel data from the algorithm's suggestion.

test.swp — used to provide executable entry points for executing the .bas script within the SolidWorks environment.

Run_CFD.java — used to conduct the STAR-CCM+ simulations and extracts data metrics of each design for mixing performance evaluation.

**Template files:**

Blank.SLDPRT — a SolidWorks file representing a micromixer without obstacles.

Design_blank.sim — a STAR-CCM+ file with pre-defined parameters for CFD simulation.

Test.xlsx — a template file used to save simulation results, convert obstacle position information, and calculate the mixing performance.
