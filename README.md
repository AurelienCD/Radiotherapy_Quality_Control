# Radiotherapy_Quality_Control
Code for radiotherapy quality control realised in typical french medical physics department

These code allows:
1) to performed the analysis of multiple dosimetric film for quality control in tomotherapy
2) to analyse dynalogs files from VARIAN VMAT radiotherapy machine. 
  a) For PFROTAT daily test : It analyses the leaf positions, and rise an error if the real position is higher of 0.5mm to the expected position
  b) For MLC-Dyn monthly test : It analyses the gantry position, gantry speed and dose rate
