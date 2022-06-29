import numpy as np
import trimesh
import os
import xlsxwriter

class STLGeometryToExcel:
    
    def __init__(self, directory):
        self.directory = directory 

    def calc_stl_geometry_data(self, single_file):
        self.mesh = trimesh.load(single_file)

        self.stl_volume = self.mesh.volume

        self.bounding_box = self.mesh.bounding_box.extents 

        self.bounding_box_volume = self.bounding_box[0] * self.bounding_box[1] *self.bounding_box[2]

        self.height = self.bounding_box[2] # Height is the z-axis
        self.x_length = self.bounding_box[0]
        self.y_length = self.bounding_box[1]

        self.bottom_area = self.x_length * self.y_length

        self.surface_area = self.mesh.area

        self.stl_file_name = os.path.basename(single_file)

        self.stl_data = [self.stl_file_name, self.stl_volume, self.bounding_box_volume, self.height, self.x_length, self.y_length, self.surface_area, self.bottom_area]

        return(self.stl_data)

    def add_excel_data_labels(self, excel_file):
        labels = ["Part Name", "Part Volume (mm^3)", "Bounding Box Volume (mm^3)", "Height (mm)", "x Length (mm)", "y Length (mm)", "Surface Area (mm^2)", "Bottom Area (mm^2)"]
        self.workbook = xlsxwriter.Workbook(excel_file)
        self.worksheet = self.workbook.add_worksheet()

        row = 0 
        column = 0
        for label in labels: 
            self.worksheet.write(row,column,label)
            column += 1 

    def write_directory_data_to_excel(self,excel_file):

        self.add_excel_data_labels(excel_file)

        working_directory = os.fsencode(self.directory)
        file_list = os.listdir(working_directory)
        
        row = 0 
        for file in file_list:
            file_name = os.fsdecode(file)
            if file_name.endswith(".stl"):
                self.calc_stl_geometry_data(self.directory + "\\" + file_name)
                row += 1 
                column = 0
                for data in self.stl_data:
                    self.worksheet.write(row,column,data)
                    column +=1
            else:
                continue

        self.workbook.close()


directory_path = "C:\CAD Files\Iso-Boundingbox Test\Screw algorithm test"
stlGeometryToExcel = STLGeometryToExcel(directory_path)

stlGeometryToExcel.write_directory_data_to_excel("ReverseEngineeringTests.xlsx")



