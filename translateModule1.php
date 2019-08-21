<?php

/**
 * Attribute VB_Name = "Module1"
 * Option Explicit
 */

/**
 * Const epsilon As Double = 0.0001
 * @var float
 */
const epsilon = 0.0001;

/**
 * Const column_offset As Long = 11
 * @var integer
 */
const column_offset = 11;

/**
 * Const width_limit As Double = 300
 * @var float
 */
const width_limit = 300;

/**
 * Const height_limit As Double = 300
 * @var float
 */
const height_limit = 300;

/**
 * Const displacement_multiplier As Double = 0.353
 * @var float
 */
const displacement_multiplier = 0.353;




/*
 * Had a look through all the rest of the functions.
 * They seem to be mostly related to the frontend side of
 * things, which in our case is MS Excel worksheets.
 *
 * Here's a list of all these functions:
 * Sub SetupConsoleWorksheet()
 * Sub SetupItemsWorksheet()
 * Sub SetupContainersWorksheet()
 * Sub SetupItemItemCompatibilityWorksheet()
 * Sub SetupContainerItemCompatibilityWorksheet()
 * Sub SetupSolutionWorksheet()
 * Sub SetupVisualizationWorksheet()
 * Sub RefreshVisualizationWorksheet()
 * Sub AnimateVisualizationWorksheet()
 * Private Sub About()
 * Private Sub SendFeedback()
 * Private Sub ResetWorkbook()
 * Sub SortItemTypes()
 * Sub SortContainerTypes()
 * Sub SetupMenuItems()
 * Private Sub WatchTutorial()
 *
 * Then there's a check like this:
 * #If Win32 Or Win64 Or (MAC_OFFICE_VERSION >= 15) Then
 * after which, it's the same functions, though renamed with
 * RibbonCall suffixes.
 *
 * Here's the list again:
 * Sub ResetWorkbookRibbonCall(control As IRibbonControl)
 * Sub SetupItemsWorksheetRibbonCall(control As IRibbonControl)
 * Sub SortItemsRibbonCall(control As IRibbonControl)
 * Sub SetupItemItemCompatibilityWorksheetRibbonCall(control As IRibbonControl)
 * Sub SetupContainersWorksheetRibbonCall(control As IRibbonControl)
 * Sub SortContainerTypesRibbonCall(control As IRibbonControl)
 * Sub SetupContainerItemCompatibilityWorksheetRibbonCall(control As IRibbonControl)
 * Sub SetupSolutionWorksheetRibbonCall(control As IRibbonControl)
 * Sub SetupVisualizationWorksheetRibbonCall(control As IRibbonControl)
 * Sub AnimateVisualizationWorksheetRibbonCall(control As IRibbonControl)
 * Private Sub SendFeedbackRibbonCall(control As IRibbonControl)
 * Private Sub WatchTutorialRibbonCall(control As IRibbonControl)
 * Private Sub AboutRibbonCall(control As IRibbonControl)
 * Sub tabActivate(ribbon As IRibbonUI)
 *
 * I'll have to look over these if there's anything relevant inside them
 * once we get all the translated logic into an actual app.
 * */



