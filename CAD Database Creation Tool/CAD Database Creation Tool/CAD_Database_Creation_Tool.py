import numpy as np
import trimesh
import os
import xlsxwriter

class STLGeometryToExcel:
    
    def __init__(self, directory):
        self.directory = directory 

    def calc_stl_bounding_box_volume(self):
        self.bounding_box = self.mesh.bounding_box.extents 
        bounding_box_volume = 1 
        for i in range(0,3): 
            bounding_box_volume = bounding_box_volume * self.bounding_box[i];
        return(bounding_box_volume)

    def calc_stl_geometry_data(self, single_file):
        self.mesh = trimesh.load(single_file)

        self.stl_volume = self.mesh.volume

        self.bounding_box_volume = self.calc_stl_bounding_box_volume()

        self.height = self.bounding_box[2] # Height is the z-axis
        self.x_length = self.bounding_box[0]
        self.y_length = self.bounding_box[1]

        self.bottom_area = self.x_length * self.y_length

        self.surface_area = self.mesh.area

        self.stl_file_name = os.path.basename(single_file)

        self.stl_data = [self.stl_file_name, self.stl_volume, self.bounding_box_volume, self.height, self.x_length, self.y_length, self.surface_area, self.bottom_area]

        return(self.stl_data)

    def add_excel_data_labels(self, excel_file):
        labels = ["part name", "part volume", "bounding box volume", "height", "x length", "y length", "surface area", "bottom area"]
        workbook = xlsxwriter.Workbook(excel_file)
        worksheet = workbook.add_worksheet()

        row = 0 
        column = 0
        for label in labels: 
            worksheet.write(row,column,label)
            column += 1 

        workbook.close()
        #workbook = xlsxwriter.Workbook(excel_file)
        #worksheet = workbook.add_worksheet()
        #worksheet.write(0,0,"beans")
        #workbook.close()

    def write_directory_data_to_excel(self,excel_file):

        self.add_excel_data_labels(excel_file)

        workbook = xlsxwriter.Workbook(excel_file)
        worksheet = workbook.add_worksheet()
        working_directory = os.fsencode(self.directory)
        file_list = os.listdir(working_directory)

        for file in file_list:
            file_name = os.fsdecode(file)
            if file_name.endswith(".stl"):
                self.calc_stl_geometry_data(self.directory + "\\" + file_name)
                
                row = file_list.index(file) 
                column = 0
                for data in self.stl_data:
                    worksheet.write(row,column,data)
                    column +=1
            else:
                continue

        workbook.close()







directory_path = "C:\CAD Files\Iso-Boundingbox Test\Screw algorithm test"
stlGeometryToExcel = STLGeometryToExcel(directory_path)

stlGeometryToExcel.write_directory_data_to_excel("ReverseEngineeringTests.xlsx")



