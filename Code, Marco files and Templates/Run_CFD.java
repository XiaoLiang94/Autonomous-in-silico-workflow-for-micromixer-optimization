package macro;

import java.util.*;

import star.common.*;
import star.base.neo.*;
import star.cadmodeler.ui.*;
import star.base.report.*;
import star.vis.*;
import star.cadmodeler.*;
import star.meshing.*;
import java.io.BufferedReader;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintWriter;

public class Run_CFD_Modified extends StarMacro {

    public void execute() {
        int numDesigns = population_number; 
        String baseSimFilePath = "D:\\Close_loop_in_silico_optimization_showcase\\Design_blank.sim";     

        for (int i = 1; i <= numDesigns; i++) {
            String simFilePath = "D:\\Close_loop_in_silico_optimization_showcase\\T_0\\Design" + i + ".sim";
            String x_tFilePath = "D:\\Close_loop_in_silico_optimization_showcase\\T_0\\Design" + i + ".x_t";
            String csvFilePath = "D:\\Close_loop_in_silico_optimization_showcase\\T_0\\Design" + i + ".csv";

            executeSimulation(baseSimFilePath, simFilePath, x_tFilePath, csvFilePath);

        try {
            Thread.sleep(10000); 
        } catch (InterruptedException e) {

            System.err.println("Thread was interrupted: " + e.getMessage());
        }
        }
    }

    private void executeSimulation(String baseSimFilePath, String simFilePath, String x_tFilePath, String csvFilePath) {
        // Load the base simulation file
        Simulation simulation = new Simulation(baseSimFilePath);

        if (simulation == null) {
            simulation.println("Failed to load simulation from " + baseSimFilePath);
            return;
        }

        // Create and initialize 3D-CAD scene
        Scene cadScene = simulation.getSceneManager().createScene("3D-CAD View");
        cadScene.initializeAndWait();

        CadModel cadModel = ((CadModel) simulation.get(SolidModelManager.class).getObject("3D-CAD Model 1"));
        simulation.get(SolidModelManager.class).editCadModel(cadModel, cadScene);
        cadScene.open();
        cadScene.setAdvancedRenderingEnabled(false);

        // Import new CAD file
        ImportCadFileFeature importCadFeature = ((ImportCadFileFeature) cadModel.getFeature("ImportCad 1"));
        ModelReImportCommands reImportCommands = importCadFeature.createModelReImportCommands();
        reImportCommands.loadFileForReplaceBody(importCadFeature, x_tFilePath);

        // Create layout views and auxiliary scenes
        LayoutView layoutView = simulation.get(LayoutViewManager.class).createLayoutView("Reimport Model", LayoutPane.Axis.LEFT_TO_RIGHT, 1, 2);
        layoutView.getLayoutAssociationManager().createAssociation(layoutView.getRootLayoutPane().getPaneAt(0), cadScene);

        Scene auxScene1 = simulation.getSceneManager().createScene("ReImport Auxiliary Scene 1");
        auxScene1.initializeAndWait();
        auxScene1.setViewLink(true);
        layoutView.getLayoutAssociationManager().createAssociation(layoutView.getRootLayoutPane().getPaneAt(1).getPaneAt(0), auxScene1);

        Scene auxScene2 = simulation.getSceneManager().createScene("ReImport Auxiliary Scene 2");
        auxScene2.initializeAndWait();
        auxScene2.setViewLink(true);
        layoutView.getLayoutAssociationManager().createAssociation(layoutView.getRootLayoutPane().getPaneAt(1).getPaneAt(1), auxScene2);

        layoutView.open();

        reImportCommands.createLayoutView(new NeoObjectVector(new Object[]{cadScene, auxScene1, auxScene2}));

        cadScene.close();

        LayoutViewUpdate layoutViewUpdate = layoutView.getLayoutViewUpdate();

        // Compare and replace bodies
        Body cadBody = ((Body) cadModel.getBody("Body 1"));
        reImportCommands.compareBodies(cadBody, "Body");

        Map<String, String> matchingFacesMap = reImportCommands.getMatchingFaces(cadBody, "Body");

        if (matchingFacesMap != null) {
            List<String> newFaceNames = new ArrayList<>();

            for (Map.Entry<String, String> entry : matchingFacesMap.entrySet()) {
                String oldFaceName = entry.getKey();
                String newFaceName = entry.getValue();
                simulation.println("Old Face: " + oldFaceName + " -> New Face: " + newFaceName);
                newFaceNames.add(newFaceName);
            }

            String[] newFacesArray = newFaceNames.toArray(new String[0]);

            Face face1 = ((Face) cadModel.getFaceByLocation(cadBody, new DoubleVector(new double[]{-0.0170392, 0.000573573, -0.00641003})));
            Face face2 = ((Face) cadModel.getFaceByLocation(cadBody, new DoubleVector(new double[]{0.0178608, 0.000573573, -0.00593376})));
            Face face3 = ((Face) cadModel.getFaceByLocation(cadBody, new DoubleVector(new double[]{-0.0170392, 0.000573573, -0.00481003})));

            reImportCommands.replaceBody(cadBody, "Body", new NeoObjectVector(new Object[]{face1, face2, face3}), new StringVector(newFacesArray));
        } else {
            simulation.println("No matching faces found.");
        }

        reImportCommands.finalizeReplaceBody();

        // Clean up layout views and scenes
        layoutView.getLayoutAssociationManager().deleteAssociationFor(layoutView.getRootLayoutPane().getPaneAt(0));
        layoutView.getLayoutAssociationManager().deleteAssociationFor(layoutView.getRootLayoutPane().getPaneAt(1).getPaneAt(0));
        layoutView.getLayoutAssociationManager().deleteAssociationFor(layoutView.getRootLayoutPane().getPaneAt(1).getPaneAt(1));

        layoutView.close();
        simulation.get(LayoutViewManager.class).removeObjects(layoutView);

        auxScene1.setViewLink(false);
        auxScene2.setViewLink(false);
        cadScene.close();

        simulation.getSceneManager().deleteScenes(new NeoObjectVector(new Object[]{auxScene1, auxScene2}));

        cadModel.update();

        simulation.get(SolidModelManager.class).endEditCadModel(cadModel);
        simulation.getSceneManager().deleteScenes(new NeoObjectVector(new Object[]{cadScene}));

        // Update parts and generate mesh
        SolidModelPart solidPart = ((SolidModelPart) simulation.get(SimulationPartManager.class).getPart("Cut-Extrude2"));
        simulation.get(SimulationPartManager.class).updateParts(new NeoObjectVector(new Object[]{solidPart}));

        MeshPipelineController meshController = simulation.get(MeshPipelineController.class);
        meshController.generateVolumeMesh();

        // Run simulation
        ResidualPlot residualPlot = ((ResidualPlot) simulation.getPlotManager().getPlot("Residuals"));
        residualPlot.open();

        simulation.getSimulationIterator().run();

        // Extract data and export to CSV
        XyzInternalTable xyzTable = ((XyzInternalTable) simulation.getTableManager().getTable("mixing index"));
        FvRepresentation fvRepresentation = ((FvRepresentation) simulation.getRepresentationManager().getObject("Volume Mesh"));
        xyzTable.setRepresentation(fvRepresentation);
        xyzTable.extract();

        try {
            xyzTable.export(csvFilePath, ",");
            simulation.println("CSV file saved successfully: " + csvFilePath);
        } catch (Exception e) {
            simulation.println("Failed to save CSV file: " + e.getMessage());
        }

        // Add pressure drop data to CSV
        ExpressionReport pressureDropReport = ((ExpressionReport) simulation.getReportManager().getReport("Pressure_drop"));
        double pressureDropValue = pressureDropReport.getReportMonitorValue();

        appendPressureDropToCsv(simulation, csvFilePath, pressureDropValue);

        // Save and close the simulation
        try {
            simulation.saveState(simFilePath);
            simulation.println("Simulation state saved successfully: " + simFilePath);
        } catch (Exception e) {
            simulation.println("Failed to save simulation state: " + e.getMessage());
        }
        simulation.close();
    }

    private Simulation loadSimulation(String simFilePath) {
        Simulation simulation = getActiveSimulation();
        if (simulation != null) {
            simulation.close();
        }

        // Load the simulation state from the specified file
        try {
            simulation = getActiveSimulation();
            simulation.loadState(simFilePath);
            simulation.println("Simulation loaded successfully from: " + simFilePath);
        } catch (Exception e) {
            if (simulation != null) {
                simulation.println("Failed to load simulation: " + e.getMessage());
            }
            simulation = null;
        }
        return simulation;
    }

    private void appendPressureDropToCsv(Simulation simulation, String csvFilePath, double pressureDropValue) {
        List<String[]> lines = new ArrayList<>();
        try (BufferedReader reader = new BufferedReader(new FileReader(csvFilePath))) {
            String line;
            while ((line = reader.readLine()) != null) {
                lines.add(line.split(","));
            }
        } catch (IOException e) {
            simulation.println("Failed to read CSV file: " + e.getMessage());
        }

        int columnIndex = 5; // Column to insert pressure drop value

        if (lines.size() > 0) {
            String[] header = lines.get(0);
            if (header.length <= columnIndex) {
                header = extendArray(header, columnIndex + 1);
            }
            header[columnIndex] = "Pressure_drop";
            lines.set(0, header);
        }

        if (lines.size() > 1) {
            String[] data = lines.get(1);
            if (data.length <= columnIndex) {
                data = extendArray(data, columnIndex + 1);
            }
            data[columnIndex] = String.valueOf(pressureDropValue);
            lines.set(1, data);
        } else {
            String[] data = new String[columnIndex + 1];
            data[columnIndex] = String.valueOf(pressureDropValue);
            lines.add(data);
        }

        try (PrintWriter writer = new PrintWriter(new FileWriter(csvFilePath))) {
            for (String[] lineArray : lines) {
                writer.println(String.join(",", lineArray));
            }
            simulation.println("CSV file updated successfully with pressure drop: " + csvFilePath);
        } catch (IOException e) {
            simulation.println("Failed to update CSV file: " + e.getMessage());
        }
    }

    private String[] extendArray(String[] original, int newLength) {
        String[] extended = new String[newLength];
        System.arraycopy(original, 0, extended, 0, original.length);
        return extended;
    }
}
