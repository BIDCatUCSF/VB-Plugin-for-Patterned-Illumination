Attribute VB_Name = "Module2"
Public GlobalVariables As Integer

    'Counter to hold number of mouse clicks
    Public theCounter As Integer
    
    'Holders for xy position
    Public xDraw(100) As Integer
    Public yDraw(100) As Integer
    
    'Holders for xy position for Multiple ROIs drawn
    Public xDrawAllROIs(1000) As Integer
    Public yDrawAllROIs(1000) As Integer
    Public IdxAllROIs(1000) As Integer
    
    'Counter for ROI vertices
    Public MasterROICounter As Integer
    
    'Number of ROI (#1-5)
    Public ROINum As Integer
    
    'holders for xy locations of tiles
    Public xTilePosGlobal(144) As Double
    Public yTilePosGlobal(144) As Double
    
    'array to tell me which tiles to fire mosaic with
    Public TilesToFire(144) As Integer
    
    'variable to adjust the brightness of the display
    Public TheBrightness As Double
    


