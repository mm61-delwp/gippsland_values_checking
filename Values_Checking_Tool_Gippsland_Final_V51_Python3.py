# --------------------------------------------------------------------------
#          Department of Energy, Environment and Climate Action
# --------------------------------------------------------------------------
#              Regional GIS Unit, Gippsland - PYTHON SCRIPT
# --------------------------------------------------------------------------
#  Name: Values checking tool
#
#  Purpose:
#   This script will cursor through individual records from a DAP layer
#   and overlay with values spatial layers.
#   The output is a line written to a table containing data to upload to QuickBase database.

#   The process will group point values for a given theme and convert points to polys.
#   This ensures that where input layers overlap, it is written into one record
#   per combination of values.
#
#   REQUIREMENTS:
#       - DAP dataset in workpath (set in code) with correctly populated fields for
#           DAP_REF_NO, District (eg "Latrobe"), Name, Schedule, Description, Risk_lvl
#       - Connection to CSDL - path to be set in code below
#       - Connection to BLDL - path to be set in code below
#       - Connection to GISdesk area of J drive
#
### NOTE: Give plenty of time for the script to run!
### Place # hashtags # at the start of the line to make it behave like a comment and bypass any coding inside it
#
# --------------------------------------------------------------------------
#  History:
#    Hans Van Elmpt - 26 Feb 2015 - Original coding started
#                   - 03 Mar 2015 - first draft completed - runs on sample data
#                   - 17 Apr 2015 - expanded to fill more fields
#                   - 02 Jun 2015 -  modified code to use state background in all input layers
#                                   to address the issue of null return on intersect
#                   - 09 Jun 2015 - Added Water, Heritage, Forest Utilization tables
#                   - 09 Jun 2015 - created a cloned version of the script for Lodon Mallee
#                   - 12 May 2016 - REVISIONS FOR 2016 DAP
#                                   - removed landscape vales
#                                   - revised Biodiversity criteria
#                   - 17 June 2016 - revised forest and forest utility table criteria
#                                   for a more robust overlay process in event of null returns
#                   - 14 June 2017 - added extra criteria for Biodiversity, Phythopthora
#                                   Updated references to FOP, WUP, TRP
#                   - 15 May 2018 - general tidy up
#                   - 19 Sep 2018 - added Cultural Heritage components
#                   - 30 Oct 2018 - made changes to rainforest and old growth source data
# Michelle Worsley  - 24 Apr 2019 - created a copy to test how script works on other computers,
#                                   .prj locations now user-defined variables
#                   - 07 Jun 2019 - Modified some data pathways to direct to more current versions, these included
#                                   Burnplan, recweb asset, land management sites and ACHRIS
#                   - 04 Jul 2019 - LRLI process draft and risk register integration
#                   - 25 Jul 2019 - Biodiversity output improved, fixing some cursors for faster processing
#                   - 17 Sep 2019 - Forest values, heritage values, and biodiversity values tidied up and ready for DMP/LRLI, with some optimisation improvements
#                   - 18 Sep 2019 - fix monitoring site values, fix heritage land manager query, commented out duplicate EVC removal, and commented out unneeded delete_management rows now that overwriteOutput is True
#                                   add TRP coupe mitigation
#                   - 19 Sep 2019 - update cursor for forest mitigations, debugging the monitoring sites and combining forest/bio values together
#                   - 30 Sep 2019 - debugging monitoring, coupe, phytoptera and trp coupe forest values so actual detections don't go skipped if the works is entirely within the values.
#                                   Sped up the 'adding geometry data' part using an update cursor
#                   - 01 Oct 2019 - moved the summarising table into heritage recursive area
#                   - 04 Oct 2019 - Cultural Heritage sites now have XY points added
#                   - 07 Oct 2019 - Cultural heritage Sites buffer changed from 200m to 550m
#                   - 09 Oct 2019 - fix up some CH, and further optimise some of the forest and biodiversity values, CH now added to values summary table
#                   - 21 Oct 2019 - further debug for CH to ensure nothing is skipped
#                   - 24 Oct 2019 - DAP component started, making another attempt for forest data optimisation with less field outputs being carried over
#                   - 29 Oct 2019 - added grazing, rainforest moved to forests section, LBP done and incorporated into risk reg
#                   - 30 Oct 2019 - intermediate features are deleted along the way to free up memory space, raretable cursor updated, debugging some forest and bio values
#                   - 31 Oct 2019 - Summary table debugging and updating Summary and Water Table cursors
#                   - 07 Nov 2019 - updated Water table to include catchment name
#                   - 14 Nov 2019 - Further debugging for Summary and Heritage themes
#                   - 18 Nov 2019 - Optimising attempt where one-off processes are done outside the recursive part of the script (LRLI values and tempvic). Fixed up some Modelled Old Growth process + mitigation
#                   - 20 Nov 2019 - EXPERIMENTAL: Substituting Identity with Spatial Join, and running one big pass through the whole data instead of going through each individual activity.
#                   - 29 Nov 2019 - LBP colony now have XY points, passes will be split up by Risk_lvl, added newer MOG data
#                   - 10 Dec 2019 - Updated Rainforest data, MOG and RFSOS removed due to no prescription for DAP/JFMP
#                   - 20 Jan 2020 - remove Murrungowar data and from the rainforest clip, because of new RAINFOR data
#                   - 23 Jan 2020 - All datasets to work with MGA z55H projection for consistency with spatial intesection checks, CH and Biodiversity debugging done, CMA added to Water,
#                                   tempvic has been buffered by 550m to accommodate activities near the district boundary.
#                   - 28 Jan 2020 - Summary now has Spatial Join instead of Identity
#                   - 29 Jan 2020 - Attempt to reduce duplicate entries in Summary and Water table, sped up CH checks
#                   - 21 Feb 2020 - Table to Excel should run automatically once the main part of script is done (saved in the specified Workpath, recommended to move and rename it to something else)
#                   - 10 Mar 2020 - removed issue with CH sites being picked up kilometres away
#                   - 13 Mar 2020 - Heritage site place name now included
#                   - 30 Mar 2020 - Risk Level field is carried through to Forest and a Biodiversity outputs, some adjustment to EVC mitigations for LRLI activities
#                   - 14 Apr 2020 - Flattening of Summary table values so one activity occupies one line
#                   - 07 May 2020 - Forest Values are now just their own spreadsheet without biodiversity or cultural heritage appended to it
#                   - 10 Jun 2020 - VBA_outputs automatically saved in workpath once script is finished running. JFMP risk register added, set mode = "JFMP" to use JFMP register instead of DAP risk register
#                   - 08 Jul 2020 - Updated JFMP Risk Register and added EC Latrobe's values to the DAP/NBFT Risk Register. Table joins have slightly been modified to accommodate EC and AGG differences.
#                                   Outputs are modified and Risk Register cleaning is entirely within an update cursor to speed up the process
#                   - 13 Jul 2020 - Modified how district list is handled, removed an additional loop - running one district at a time is still recommended. Fixed some issues with Risk Register.
#                   - 04 Aug 2020 - District info handled slightly differently so shapefiles that cover more than one district can be run without issues. ***Biodiversity theme should still be run under the same District only.***
#                   - 15 Dec 2020 - Modified how Monitoring sites are handled, included contact details and prescriptions in the description fields
#                   - 4 Jan 2021 - Extra code added for JFMP heritage checks, mainly for Contingecy burn areas
#                   - 6 Jan 2021 - additional dissolve for DAP_buff features to prevent self intersections
#                   - 16 Feb 2021 - Adjusted biodiversity output table so nothing gets deleted in the filter - leaving in "no comment..." values advice
#        V25        - 10 Mar 2021 - Changed CH preliminary data source to gigis N drive and new data + field for joint managed parks and reserves within GLaWAC. Fire_Sensitivity also carried over into heritage output
#        V26        - 23 Mar 2021 - Water theme changed from PWSC to using HY_WATERCOURSE to determine works on waterways
#        V27        - 28 Apr 2021 - CH theme modified to create 2 separate outputs
#        V28        - 18 May 2021 - CH changes to LandInfo, PV district data added, handled Risk Registers differently (DAP risk reg no longer uses Risk_Landscape), xls outputs to be converted to csv, and additional output for works details(.gdb input in table form)
#        V29        - 24 May 2021 - Changed how ACH sites and prelims reports are merged, Biodivesity now also avoid duplicate XYs for the same species, adding Quickbase key fields at the end (QB_ID)
#        V30        - 07 Jun 2021 - Two new outputs made specific for QuickBase, combined versions for biodiversity + forests and heritage site + land info
#        V31        - 18 Jun 2021 - Land info for heritage now part of broader Works detail (added to works_features) - new biodiversity CAMS data added to rarelist species to expand the search for more taxons. SRO Values check added.
#        V32        - 16 Sep 2021 - (not draft version) VBA restricted datasets, flora 25 and ARI_AQUA_CATCH added to biodiversity output. Victoria heritage register and Victorian heritage inventory also added to forests module
#        V33        - 22 Sep 2021 - (not draft version) CH fixing and removed the now obsolete union and filtering processes to greatly speed up the CH checking process
#        V34        - 28 Sep 2021 - EXTRA_INFO flattened in biodiversity theme to avoid QB_ID duplicates
#        V35        - 15 Nov 2021 - Biodiversity mitigation table cleaning debugged
#        V36        - 04 Jan 2022 - SCRIPT_DATE now uses MMDDYYYY format for more seamless importing into Quickbase
#        V37        - 09 May 2022 - change CH key QBID generation, and added railways and utilities to the values checking
#        V38        - 18 May 2022 - applied 100m buff for SRO checks, added new JFMP contingency buffers for biodiversity values, BUFFER_TYPE added to biodiversity template output, additional biodiveristy output for JFMP (one for JFMP Area, and another for Contingency Area)
#        V39        - 27 Jun 2022 - Updated _WorksDetail output spreadsheet to use a designated template to continue matching the pipeline configuration while using a new DAP template. JFMP now outputs their shapefiles with 500m buffers.
#                                   Grazing licenses now include Tenure ID numbers in their description
#        V40        - 16 Mar 2023 - Updated Apiary mitigation to include more detail
#        V41        - 06 Jun 2023 - TEST new Biodiversity fields to be added to Biodiv_summary_rawdata and DAP_Biodiversity_Forests_combinedQB, updates to the Value ID for Biodiversity (currrently specific to JFMP and NBFT)
#                   - 19 Jun 2023 - Updated templates for JFMP and NBFT, Pest plant data source has been modified, some temporary data locations while the CSDL is being updated
#        V42        - 28 Jun 2023 - Updates to Forest values with improved QB_IDs with the inclusion of a specific VALUE_ID to identify specific values, these will be set as QB_ID2 in the Forests_summary table output (not the Bio+Forests QB one).
#                                   Inclusion of Native Title requirements for DAP
#        V43        - 31 Jul 2023 - Some optimization for Phytophtora risk polygon
#        V44        - 10 Aug 2023 - Temp measure: with duplicates still appearing with the transition of QB_ID and additional fields in _BiodiversityValues and _Forest_summary tables, another round of cleanup to remove duplicates has been added
#        V45        - 27 Sep 2023 - Adjusted some wording for low impact/exempt activity NT assessment status
#        V46.5      - 12 Feb 2024 - Move Rainforest to Biodiversity theme, Joint Managed Park now uses a new datasource (PLM25_OVERLAYS_TO)
#        V47        - 27 May 2024 - Changed DAP and contingency buffering to remove true curves and improve spatial accuracy
#        V48        - 11 Jun 2024 - DAP and NBFT use the same bio+forests output to make up for new VBA fields to add to quickbase, catchment has changed from one to many to a one to one spatial join to speed up process
#        V49        - 23 Aug 2024 - Update on temporary measure for removing value duplicates after mitigation and QB_IDs are calculated for the biodivesity output (still some more work to do to clear out redundant mitigations for EVC)
#        V50        - 20 Nov 2024 - Script now compatible with Python 3 and (test) direct imports to VMA Quickbase if work is DAP or Ad Hoc (spoiler alert, we tried the direct import test and it didn't work, oh well)
#        V50.1      - 04 Feb 2025 - Minor edit: Value ID for forest value now carry into bio+forest output field RECORD_ID
#        V51        - 08 Apr 2025 - Slightly updated biodiversity output EVC filtering (removing duplicate "no comment..." records for the same EVC names that do have other mitigations attached to the same job),
#                                   updated JFMP biodiversity outputs so the correct Risk Register is being used for both Burn Unit and Contingency Area, Works detail output for JFMP and NBFT have ASSET_ID renamed to FMS_ID
# ==========================================================================

import arcpy, os, string, datetime, glob
import pandas as pd
from arcpy import env
import json
import requests
# from google.oauth2 import service_account
# from google.auth.transport.requests import Request



###############################################################################################################
#                         USER-DEFINED VARIABLES - Set the values as required                                 # 
###############################################################################################################

# processing mode
mode = "JFMP"  # Options: "DAP", "JFMP", "NBFT" - Any for ad-hoc

# list of themes to process
theme_list = ["forests", "biodiversity"] # "summary", "forests", "biodiversity", "heritage", "water"

# output folder
workpath = r"C:\data\dap_temp"  # root path of workspace for outputs and data - must have several GB free space

# data locations
CSDL = "D:\\Data\\gis_public\\CSDL"  # Path of the CSDL source folder
CSDL2 = "D:\\Data\\gis_public\\CSDL FloraFauna2"  # Path of the CSDL-restricted source folders
CSDL3 = "D:\\Data\\gis_public\\CSDL Culture"  # Path of the CSDL-restricted cultural source folders
BLD = "O:\\regional_business.gdb"  # Path for business level data library
REG = "D:\\Data\\gis_public\\gisdesk\\GISData"  # Regional replicated data source eg Pest Animals
works_features = r"C:\Users\mw1m\OneDrive - Department of Energy, Environment and Climate Action\GIS\gigis Projects\DAP\Data\QA.gdb\works_shapefil_ExportFeature"  # r"C:\Data\GIS_Projects_Local\DAP\Data\QA.gdb\JFMP_TEST_Bio"

# options
output_shapefile = False  # set to True if works need to be exported and saved individually as shapefiles in another folder, these later get saved in an ECM folder for sharing
shapefilepath = workpath[:-4] + "QB_Processing\\Shapefiles"  # folder location to save the shapefiles


### For Risk Register:
DAPriskreg = workpath + r"\RiskRegister.gdb\NBFTDAP_RiskRegister"  # use this for DAP or NBFT, full risk register with all EVCs
###LRLIriskreg = "C://Data//GIS_Projects_Local//DAP//Data//RiskRegister.gdb//LRLI_RiskRegister"  # filtered out values that won't be threatened under LRLI (additional EVCs removed)
JFMPriskreg = workpath + r"\RiskRegister.gdb\JFMP_RiskRegister"  # use this for FOP/JFMP only - combined advice for both EC and AGG BRL


###############################################################################################################
#                                SCRIPT VARIABLES - DO NOT CHANGE THESE!                                      # 
###############################################################################################################

outGDB = "DAP_Checking_outputs.gdb"  # Name of output GDB
templateGDB = os.path.join(workpath, "\\DAP_table_templates.gdb")  # Name of table template GDB within workspace

# Summary table inputs
in_VIC = os.path.join(CSDL, "FIRE.GDB", "LF_DISTRICT")
in_FMZ = os.path.join(CSDL, "FORESTS.GDB", "FMZ100")
in_PARCEL = os.path.join(CSDL, "VMPROP.GDB", "PARCEL_CROWN_APPROVED")
in_PLM = os.path.join(CSDL, "CROWNLAND.GDB", "PLM25")
in_PLMOVRLAY = os.path.join(CSDL, "CROWNLAND.GDB", "PLM25_OVERLAYS")
in_FTYPE = os.path.join(CSDL, "FORESTS.GDB", "FORTYPE500")
in_PLAN = os.path.join(CSDL, "VMPLAN.GDB", "PLAN_ZONE")
in_PLOLAY = os.path.join(CSDL, "VMPLAN.GDB", "PLAN_OVERLAY")
in_NT = os.path.join(REG, "RegionalData.gdb", "GUNAIKURNAI_DETERMINATION")

# Water inputs
in_HYDRO = os.path.join(CSDL, "VMHYDRO.gdb", "HY_WATERCOURSE")
in_CMA = os.path.join(CSDL, "CATCHMENTS.gdb", "CMA100")

# Biodiversity inputs
in_RARETABLE = os.path.join(workpath, "Biodiversity", "VBA_EXTRACTS.gdb", "Victorian_FMP_And_CAMS")
in_SRO = os.path.join(REG, "RegionalData.gdb", "BLD_SpeciesRecoveryOverlay_20220410")
in_VBAFL25 = os.path.join(CSDL, "FLORAFAUNA1.GDB", "VBA_FLORA25")
in_VBAFLTHR = os.path.join(CSDL2, "FLORAFAUNA2.GDB", "VBA_FLORA_THREATENED")
in_VBAFLRES = os.path.join(CSDL2, "FLORAFAUNA2.GDB", "VBA_FLORA_RESTRICTED")
in_VBAFA25 = os.path.join(CSDL, "FLORAFAUNA1.GDB", "VBA_FAUNA25")
in_VBAFATHR = os.path.join(CSDL2, "FLORAFAUNA2.GDB", "VBA_FAUNA_THREATENED")
in_VBAFARES = os.path.join(CSDL2, "FLORAFAUNA2.GDB", "VBA_FAUNA_RESTRICTED")
in_VBATaxaList = os.path.join(CSDL, "FLORAFAUNA1.GDB", "VBA_TAXA_LIST")
in_AQUACATCH = os.path.join(CSDL, "FLORAFAUNA1.GDB", "ARI_AQUA_CATCH")
in_LBP = os.path.join(CSDL, "FLORAFAUNA1.GDB", "LBPAG_BUFF_CHRFA")
in_EVC = os.path.join(CSDL, "FLORAFAUNA1.GDB", "NV2005_EVCBCS")
in_RFCLIP = os.path.join(workpath, "Misc_data.gdb", "CH_RAINFOREST_CLIP")
in_RFPOLY = os.path.join(CSDL, "FORESTS.GDB", "RAINFOR")
in_RFPOLYCH = os.path.join(CSDL, "FORESTS.GDB", "RAINFOR100_CH")

# Forest management inputs
in_HUTS = os.path.join(CSDL, "FORESTS.GDB", "EG_ALPINE_HUT_SURVEY")
in_RECWEBS = os.path.join(CSDL, "FORESTS.GDB", "RECWEB_SITE")
in_RECWEBA = os.path.join(CSDL, "FORESTS.GDB", "RECWEB_ASSET")
in_RECWEBH = os.path.join(CSDL, "FORESTS.GDB", "RECWEB_HISTORIC_RELIC")
in_RECWEBN = os.path.join(CSDL, "FORESTS.GDB", "RECWEB_SIGN")
in_RECWEBC = os.path.join(CSDL, "FORESTS.GDB", "RECWEB_CARPARK")
in_MONITOR = os.path.join(REG, "RegionalData.gdb", "BLD_LAND_MANAGEMENT_SITES")
in_GTREES = os.path.join(CSDL, "FORESTS.GDB", "EG_GIANT_TREES")
in_PESTPL = os.path.join(REG, "RestrictedData.gdb", "MAX_EIP_HRIP_Open_plm25_20230614")
in_TRP = os.path.join(CSDL, "FORESTS.GDB", "TRP")
in_BURN = os.path.join(CSDL, "FIRE.GDB", "BURNPLAN25")
in_MINSITE = os.path.join(CSDL, "MINERALS.GDB", "MINSITE")
in_MIN = os.path.join(CSDL, "MINERALS.GDB", "MIN")
in_APIARY = os.path.join(CSDL, "CROWNLAND.GDB", "APIARY_BUFF")
in_GLIC = os.path.join(CSDL, "VMCLTENURE.GDB", "V_CL_TENURE_POLYGON_DC")
in_TCODE = os.path.join(CSDL, "VMREFTAB.GDB", "CL_TENURE_DESC")
in_CHEM = os.path.join(workpath, "Misc_data.gdb", "BLD_CHEMICAL_CONTROL_AREAS")
in_PCRISK = os.path.join(workpath, "Misc_data.gdb", "HIGH_PC_RISK_0610_POLY")

# Utilities inputs
in_POWRLINE = os.path.join(CSDL, "VMFEAT.gdb", "POWER_LINE")
in_PIPELINE = os.path.join(CSDL, "VMFEAT.gdb", "FOI_LINE")
in_RAIL = os.path.join(CSDL, "VMTRANS.gdb", "TR_RAIL")

# Historic sites inputs
in_HIST100 = os.path.join(CSDL, "FLORAFAUNA1.GDB", "HIST100_POINT")
in_VHI = os.path.join(CSDL, "PLANNING.gdb", "HERITAGE_INVENTORY")
in_VHR = os.path.join(CSDL, "PLANNING.gdb", "HERITAGE_REGISTER")

# Cultural heritage inputs
in_ACHRISS = os.path.join(CSDL3, "CULTURE.gdb", "ACHP_FIRESENS")
in_ACHRISP = r"N:\projects\prelim_ch_change_detection\data\data.gdb\CH_PRELIMINARY_REPORTS_ALL"
in_RAP = os.path.join(CSDL, "CULTURE.gdb", "RAP")
in_SENS = os.path.join(CSDL, "CULTURE.gdb", "SENSITIVITY_PUBLIC")
in_JOINTMGMT = os.path.join(CSDL, "CROWNLAND.GDB", "PLM25_OVERLAYS_TO")
in_PV = os.path.join(CSDL, "FIRE.GDB", "PV_DISTRICTS")


# Output table variables
sumtab = "DAP_Summary"
biotab = "DAP_Biodiv_summary_rawdata2"
wattab = "DAP_Water_summary"
hertab1 = "DAP_Heritage_summary_SitesInfo"
hertab2 = "DAP_Heritage_summary_LandInfo"
fortab = "DAP_Forest_summary2"
biofmQB = "DAP_Biodiversity_Forests_combinedQB"  # Older version, for DAP and Ad hoc
biofmQB2 = "DAP_Biodiversity_Forests_combinedQB_2"  # for JFMP/NBFT
hertabQB = "DAP_Heritage_SiteInfoQB"
worksfc = "works_shapefile_Template"  # older template that still has all original fields which matches with connected tables in VMA
worksfc2 = "Works_Shapefile_Template2"  # For JFMP/NBFT
worksdetailog = "DAP_WorksDetail_QB"  # original worksdetail fields in tabular form

# spatial reference systems
sr_gda_z55 = arcpy.SpatialReference(28355)  # GDA94 MGA Zone 55
sr_vicgrid = arcpy.SpatialReference(3111) # VICGRID94


###############################################################################################################
#                                     EXECUTABLE PART OF THE SCRIPT                                           # 
###############################################################################################################

# set up geoprocessing environment
arcpy.env.overwriteOutput = True
env.workspace = os.path.join(workpath, outGDB)
arcpy.env.outputCoordinateSystem = sr_gda_z55 # GDA94 MGA Zone 55
arcpy.env.cartographicCoordinateSystem = sr_gda_z55 # GDA94 MGA Zone 55
arcpy.env.overwriteOutput = True
arcpy.env.parallelProcessingFactor = "75%"
arcpy.SetLogHistory(False)


# populate time and date variables
now = datetime.datetime.now()
start_time = now.strftime("%H:%M:%S")
start_date = now.strftime("%d%m%Y")

print("Running the Values Checking tool")
print(f"Start time is {start_time} on {now.strftime('%A %d %B %Y')}")

# make file gdb for storage
if arcpy.Exists(outGDB):
    print(f"Deleting {outGDB}")
    arcpy.management.Delete(outGDB)

print("Creating a new output GDB")
arcpy.management.CreateFileGDB(workpath, outGDB)

# create a local copy of the works and values layers
print("Creating a copy of the works layer from original supplied")
arcpy.management.CopyFeatures(works_features, "works_features", sr_gda_z55, f"{reference_field} <> ''", )

# add geometry details to works
print("\nAdding geometry data")

# add geometry fields if they don't exist
field_list = arcpy.ListFields("works_features")
for new_field in ["Easting", "Northing", "Length_Km", "Area_Ha"]:
    if new_field not in [field_list]:
        arcpy.management.AddField("works_features", f"new_field", "DOUBLE", "7")

with arcpy.da.UpdateCursor("works_features", ["SHAPE@", "Easting", "Northing", "Length_Km", "Area_Ha"]) as cursor:
    for row in cursor:
        (shape, easting, northing, len_km, area_ha) = row
        easting = int(shape.x)
        northing = int(shape.y)
        len_km = (shape.length / 1000) * 0.498  # approximate centreline length for roading lines
        area_ha = shape.area / 10000
        cursor.updateRow(row)


# Make a blank, editable copy of tables from templates
# use different bio+forest??? and worksdetail output if JFMP or NBFT are being run (slightly different fields)
opt_works = worksfc2 if 'JFMP' in mode or 'NBFT' in mode else opt_works = worksfc
for tab in [sumtab, biotab, wattab, hertab1, hertab2, fortab, hertabQB, opt_works, biofmQB]:
    arcpy.management.Copy(os.path.join(templateGDB, tab), tab)


def add_and_fill_text_field(input, field_name, type, str):
    """ Re-usable function to add and populate a field """
    arcpy.management.AddField(f"{input}", "f{field_name}", "TEXT", "", "", 50)
    arcpy.management.CalculateField(f"{input}", "f{field_name}", f"{str}", "PYTHON3")

def add_xy(input):
    """ Re-usable function to add and populate X and Y geometry fields """
    
    try:
        # add coordinates fields
        arcpy.management.AddField("{input}", "X", "DOUBLE", "", "0", "10")
        arcpy.management.AddField("{input}", "Y", "DOUBLE", "", "0", "10")

        # populate coordinates from SHAPE@
        with arcpy.da.UpdateCursor("temppestv", ['SHAPE@', "X", "Y"]) as cursor: 
            
            # alias fields for sanity
            (shape, easting, northing) = row
            
            # step through rows and populate X and Y
            for row in cursor:
                easting = int(shape.x)
                northing = int(shape.y)
                cursor.updateRow(row)
    
    except Exception as e:
        print(f"Error executing add_xy: {str(e)}") 


# Generate  a list of targetted rare species from in_RARETABLE to use in selection loops
with arcpy.da.SearchCursor(in_RARETABLE, "TAXON_CODE", "TAXON_CODE > 0") as cursor:
    rarelist = ",".join(str(row[0]) for row in cursor)
rarelist_query = f"(STARTDATE > date '1980-01-01 00:00:00') AND MAX_ACC_KM <= 0.5 AND TAXON_ID in ({rarelist})" #this query will be used later in biodiversity querying for flora records

# setting up one-off features to use for finding values, depending on theme - only LRLI values run here because it applies universally to all activities
# only needs to be done once, which is why it's here and not inside the recursive part of the script
if "forests" in theme_list:

    #### PESTS
    print("   Creating one-off Pests data layer...")

    # create buffered version of values layer
    arcpy.analysis.Buffer(in_PESTPL, "temppestv", "10 meter", "FULL", "ROUND", "LIST", ["Species", "Ranking"])

    # add coordinates
    add_xy("temppesttv")

    # add and populate Value_Type field
    add_and_fill_text_field("temppestv", "Value_Type", "PEST_PLANT")

    # combine all pest species into single Value field
    arcpy.management.AddField("temppestv", "Value", "TEXT", "", "", 250)

    with arcpy.da.UpdateCursor("temppestv", ["Species", "Ranking", "Value"]) as cursor:  

        # alias fields for sanity
        (species, ranking, value) = row

        # step through rows and update
        for row in cursor:
            value = f"{species}, {ranking} Weed"
            cursor.updateRow(row)

    #### MONITORING SITES
    print("   Creating one-off forest and fire monitoring site data...")

    # field mapping for consistent value identifiers and description
    field_map = arcpy.FieldMappings()
    field_map.addTable(in_MONITOR)
    field_map.findFieldMapIndex("SITE_ID").outputField.name = "Value"
    field_map.findFieldMapIndex("SITE_NAME").outputField.name = "Value_Description"
    field_map.findFieldMapIndex("LMS_ID").outputField.name = "Value_ID"

    # create local values layer with selection
    arcpy.conversion.FeatureClassToFeatureClass(in_MONITOR, os.path.join(workpath, outGDB), "monforestfire",
                                                "(SITE_CATEGORY IN('FOREST','FIRE')) AND TIMEFRAME <> 'NOT ACTIVE'",
                                                field_mapping=field_map)
    
    # add coordinates
    add_xy("temppesttv")

    # add and populate Value_Type field
    add_and_fill_text_field("monforestfire", "Value_Type", "Monitoring Site")


    #### PHYTOPTHORA
    print("   Creating one-off High risk Phytopthora sites data...")

    # field mapping for consistent value identifiers and description
    field_map = arcpy.FieldMappings()
    field_map.addTable(in_PCRISK)
    field_map.findFieldMapIndex("CLASS").outputField.name = "Value"

    # create local dissolved values layer with selection
    arcpy.conversion.FeatureClassToFeatureClass(in_PCRISK, "in_memory", "temp_pc", "CLASS = 'High'", field_mapping=field_map)
    arcpy.management.Dissolve("in_memory\\temp_pc", os.path.join(workpath, outGDB, "pchighrisk"), 
                              ["Value"], None, "MULTI_PART", "DISSOLVE_LINES")

    # add and populate Value_Type field
    add_and_fill_text_field("monforestfire", "Value_Type", "Phytophthora Risk")
    

    #### CHEMICAL CONTROL AREAS
    print("   Creating temporary Chemical Control Area data...")

    # field mapping for consistent value identifiers and description
    field_map = arcpy.FieldMappings()
    field_map.addTable(in_PCRISK)
    field_map.findFieldMapIndex("ACCA_NAME").outputField.name = "Value"

    # create local values layer
    arcpy.conversion.FeatureClassToFeatureClass(in_CHEM, os.path.join(workpath, outGDB), "acca", field_mapping=field_map)

    # add and populate Value_Type field
    add_and_fill_text_field("acca", "Value_Type", "Agricultural Chemical Control Area")


    #### HISTORIC HERITAGE
    print("   Creating one-off statewide Historic Heritage data...")

    # field mapping for consistent value identifiers and description
    vhi_field_map = arcpy.FieldMappings()
    vhi_field_map.addTable(in_VHI)
    vhi_field_map.findFieldMapIndex("VHI_NUM").outputField.name = "Value"
    vhi_field_map.findFieldMapIndex("SITE_NAME").outputField.name = "Value_Description" 
    vhi_field_map.findFieldMapIndex("HERMES_NUM").outputField.name = "Value_ID"

    vhr_field_map = arcpy.FieldMappings()
    vhr_field_map.addTable(in_VHR)
    vhr_field_map.findFieldMapIndex("VHR_NUM").outputField.name = "Value"
    vhr_field_map.findFieldMapIndex("SITE_NAME").outputField.name = "Value_Description"
    vhr_field_map.findFieldMapIndex("HERMES_NUM").outputField.name = "Value_ID"

    hist_field_map = arcpy.FieldMappings()
    hist_field_map.addTable(in_HIST100)
    hist_field_map.findFieldMapIndex("NAME").outputField.name = "Value"

    # create buffered and vanilla versions of values layers
    arcpy.analysis.Buffer(in_HIST100, "HIST100_BUFF", "250 meter", "FULL", "ROUND", "NONE", field_mapping=hist_field_map)
    arcpy.conversion.FeatureClassToFeatureClass(in_VHI, os.path.join(workpath, outGDB), "VHI", field_mapping=vhi_field_map)
    arcpy.conversion.FeatureClassToFeatureClass(in_VHR, os.path.join(workpath, outGDB), "VHR", field_mapping=vhr_field_map)

    # clean rogue newlines from hist100
    arcpy.management.CalculateField("HIST100_BUFF", "Value", '" ".join(!Value!.split())',"PYTHON3")  # to fix names with rogue newlines

    # merge layers
    arcpy.management.Merge(["HIST100_BUFF", "VHI", "VHR"], "HistHeritage")
    add_and_fill_text_field("HistHeritage", "Value_Type", "Historic Heritage Site")
    
    # add x and y coordinates
    add_xy("HistHeritage")

    # delete temporary feature classes
    for fc in ["HIST100_BUFF", "VHI", "VHR"]:
        arcpy.management.Delete(fc)

    #### MINING SITE
    print("   Creating one-off Mining site data...")

    # field mapping for consistent value identifiers and description
    field_map = arcpy.FieldMappings()
    field_map.addTable(in_PCRISK)
    field_map.findFieldMapIndex("MINE_NAME").outputField.name = "Value"
    field_map.findFieldMapIndex("EASTING").outputField.name = "X"
    field_map.findFieldMapIndex("NORTHING").outputField.name = "Y"
    field_map.findFieldMapIndex("SITEID").outputField.name = "Value_ID"

    # create buffered version of values layer
    arcpy.analysis.Buffer(in_MINSITE, "buffmsit", "20 meter", "FULL", "ROUND", "NONE", field_mapping=field_map)

    # add and populate Value_Type field
    add_and_fill_text_field("buffmsit", "Value_Type", "Mining Site")


    #### REGIONAL JFMP DATA
    print("   Creating one-off regional JFMP data...")

    # field mapping for consistent value identifiers and description
    field_map = arcpy.FieldMappings()
    field_map.addTable(in_BURN)
    field_map.findFieldMapIndex("TREAT_NO").outputField.name = "Value"
    field_map.findFieldMapIndex("TREAT_NAME").outputField.name = "Value_Description"

    # create local values layer
    arcpy.conversion.FeatureClassToFeatureClass(in_BURN, os.path.join(workpath, outGDB), "FOP_Gipps", "REGION = 'Gippsland'",
                                                 field_mapping=field_map)
    
    # add and populate Value_Type field
    add_and_fill_text_field("FOP_Gipps", "Value_Type", "Joint Fuel Management Plan Site")


    #### TRP DATA
    print("   Creating one-off TRP data...")

    # field mapping for consistent value identifiers and description
    field_map = arcpy.FieldMappings()
    field_map.addTable(in_TRP)
    field_map.findFieldMapIndex("COUPE").outputField.name = "Value"

    # create local values layer
    arcpy.conversion.FeatureClassToFeatureClass(in_TRP, os.path.join(workpath, outGDB), field_mapping=field_map)

    # add and populate Value_Type field
    add_and_fill_text_field("TRP", "Value_Type", "TRP Coupe")


    #### UTILITIES DATA
    print("   Creating one-off Utilities data...")

    # field mapping for consistent value identifiers and description
    power_field_map = arcpy.FieldMappings()
    power_field_map.addTable(in_POWRLINE)
    power_field_map.findFieldMapIndex("FEATURE_TYPE").outputField.name = "Value"

    pipe_field_map = arcpy.FieldMappings()
    pipe_field_map.addTable(in_PIPELINE)
    pipe_field_map.findFieldMapIndex("NAME_LABEL").outputField.name = "Value"
    pipe_field_map.findFieldMapIndex("FEATURE_SUBTYPE").outputField.name = "Value_Description"

    # create temp values layers
    arcpy.conversion.FeatureClassToFeatureClass(in_PIPELINE, "in_memory", "temp_pipe", "PIPELINE", "FEATURE_TYPE = 'pipeline'", field_mapping=pipe_field_map)
    arcpy.conversion.FeatureClassToFeatureClass(in_POWRLINE, "in_memory", "temp_powr", field_mapping=power_field_map)

    # add and populate missing/aggregate fields
    add_and_fill_text_field("in_memory\\temp_powr", "Value_Description", "!FEATURE_SUBTYPE! + ' ' + !VOLTAGE!")
    add_and_fill_text_field("in_memory\\temp_powr", "Value_Type", "Powerline")
    add_and_fill_text_field("in_memory\\temp_pipe", "Value_Type", "Pipeline")

    # merge temporary leyers to gdb
    arcpy.management.Merge(["in_memory\\temp_pipe", "in_memory\\temp_powr"], "Utilities")


    #### RAILWAY
    print("   Creating one-off Railway data...")

    # field mapping for consistent value identifiers and description
    field_map = arcpy.FieldMappings()
    field_map.addTable(in_RAIL)
    field_map.findFieldMapIndex("NAME").outputField.name = "Value"
    field_map.findFieldMapIndex("FEATURE_TYPE_CODE").outputField.name = "Value_Description"

    # create local values layer
    arcpy.conversion.FeatureClassToFeatureClass(in_RAIL, os.path.join(workpath, outGDB), "RAIL", field_mapping=field_map)

    # Rename empty values 
    arcpy.management.CalculateField("RAIL", "Value", '"Unnamed" if !Value! is None else !Value!', "PYTHON3")

    # add and populate Value_Type field
    add_and_fill_text_field("RAIL", "Value_Type", "TRP Railway")


    #### APIARY
    print("   Creating one-off Apiary layer...")

    # field mapping for consistent value identifiers and description
    field_map = arcpy.FieldMappings()
    field_map.addTable(in_APIARY)
    field_map.findFieldMapIndex("TENURE_ID").outputField.name = "Value"
    field_map.findFieldMapIndex("FEATURE_TYPE_CODE").outputField.name = "Value_Description"

    # create local values layer
    arcpy.management.Copy(in_APIARY, "bee")
    arcpy.conversion.FeatureClassToFeatureClass(in_APIARY, os.path.join(workpath, outGDB), "bee", field_mapping=field_map)

    # add coordinates
    add_xy("bee")

    # add and populate Value_Type field
    add_and_fill_text_field("bee", "Value_Type", "Apiary Site")

if "biodiversity" in theme_list:

    print("   Setting up VBA template...")
    arcpy.management.Copy(templateGDB + "\\VBA_template_z55_V5", "VBA_outputs")


    #### BIODIVERSITY MONITORING
    print("   Creating one-off biodivesity monitoring data...")

    # field mapping for consistent value identifiers and description
    field_map = arcpy.FieldMappings()
    field_map.addTable(in_MONITOR)
    field_map.findFieldMapIndex("LMS_ID").outputField.name = "RECORD_ID"

    # create local values layer
    filter = "(SITE_CATEGORY IN('FAUNA', 'FLORA') OR RISK_CODE = 'LittoralRF') AND TIMEFRAME <> 'NOT ACTIVE'"
    arcpy.conversion.FeatureClassToFeatureClass(in_MONITOR, os.path.join(workpath, outGDB), "monflorafauna", field_mapping=field_map)

    # add and populate various fields
    add_and_fill_text_field("monflorafauna", "SCI_NAME", '"str(!SITE_NAME!) + " | " + str(!SITE_PRESCRIPTIONS![0:150]) + " | " + str(!CONTACT!)"')
    add_and_fill_text_field("monflorafauna", "COMM_NAME", "!RISK_CODE!")
    add_and_fill_text_field("monflorafauna", "TYPE", "Bio Monitoring")

    # add coordinates
    add_xy("monflorafauna")


# ===================================================================================================================================
#     LOOPS TO CREATE SEARCHABLE THEME-BASED OVERLAY OUTPUTS BY DAP ACTIVITY
# ===================================================================================================================================



###########
###########     PASS 1 - UP TO HERE
###########


# field references
reference_field = "DAP_REF_NO"  # set selection field name for ref number #DAP_REF_NO
district_field = "DISTRICT"  # set selection field name for district
risk_field = "RISK_LVL"  # set selection field name for Risk (distinguishes between DAP or DMP/LRLI)
DAPfields = ["DAP_REF_NO", "DAP_NAME", "SCHEDULE", "RISK_LVL", "DESCRIPTION", "AREA_HA", "LENGTH_KM"]

# determine district name from first polygon in works shapefile or feature class
with arcpy.da.SearchCursor(works_features, district_field) as cursor:
    district_name = next(cursor)[0]

print("\nProcessing records for {district_name}")

## Create a background poly layer for state of Victoria or District. This is used on layers with non-contiguous features to fix an issue with intersect
##arcpy.conversion.FeatureClassToFeatureClass(in_VIC, workpath + "\\" + outGDB, "tempvic", "\"STATE\" = 'VIC' AND \"FEATURE_TYPE_CODE\" = 'mainland'" )
##arcpy.conversion.FeatureClassToFeatureClass(in_VIC, workpath + "\\" + outGDB, "tempvic", "DISTRICT_NAME = UPPER('" + dist + "')") #to narrow down on District only
arcpy.conversion.FeatureClassToFeatureClass(in_VIC, workpath + "\\" + outGDB, "tempvic")  # "REGION_NAME = 'GIPPSLAND'"
arcpy.analysis.Buffer("tempvic", "tempvic_buff", "550 meters",
                      dissolve_option="ALL")  # buffered to increase clip area, so anything just outside the border doesn't get left out.
arcpy.management.Delete("tempvic")
arcpy.management.Rename("tempvic_buff", "tempvic")

RiskValues = set()
with arcpy.da.SearchCursor(works_features, risk_field) as cur:
    for row in cur:
        if row[0] not in RiskValues:
            RiskValues.add(row[0])

for RiskLevel in RiskValues:

    # Create a dataset for a single record from local DAP - to be used in overlay process
    print("")

    arcpy.conversion.FeatureClassToFeatureClass("works_features", workpath + "\\" + outGDB, "DAP_temp",
                                                '"' + risk_field + '" = ' + "'" + RiskLevel + "'")
    numberoffeatures = arcpy.management.GetCount("DAP_temp")
    arcpy.management.AddField("DAP_temp", "BUFFER_TYPE", "TEXT", 100)

    # Buffering the DAP shape
    # print " PROCESSING " + str(featurecount) + " of " + str(numberoffeatures) + " for " + DAP_id + " - started " + (datetime.datetime.now().strftime("%H:%M:%S"))
    print(" PROCESSING " + str(numberoffeatures) + " FEATURES for risk level: " + str(RiskLevel) + " - started " + (
        datetime.datetime.now().strftime("%H:%M:%S")))
    arcpy.analysis.Buffer("DAP_temp", "DAP_buff1",
                          "1 meters")  # dissolve_option="LIST", dissolve_field=["DAP_REF_NO", "DAP_NAME", "DISTRICT", "SCHEDULE", "RISK_LVL", "DESCRIPTION", "AREA_HA", "LENGTH_KM", "Easting", "Northing", "BUFFER_TYPE"])
    arcpy.edit.Densify("DAP_buff1", "ANGLE", "", "", "0.10")
    
    arcpy.analysis.Buffer("DAP_temp", "DAP_buff50", "50 meters", dissolve_option="LIST",
                          dissolve_field=["DAP_REF_NO", "DAP_NAME", "DISTRICT", "SCHEDULE", "RISK_LVL", "DESCRIPTION",
                                          "Easting", "Northing", "YEAR_WORKS", "AREA_HA", "LENGTH_KM", "Easting",
                                          "Northing", "BUFFER_TYPE"])
    arcpy.edit.Densify("DAP_buff50", "ANGLE", "", "", "0.10")
    arcpy.analysis.Buffer("DAP_temp", "DAP_buff100", "100 meters", dissolve_option="LIST",
                          dissolve_field=["DAP_REF_NO", "DAP_NAME", "DISTRICT", "SCHEDULE", "RISK_LVL", "DESCRIPTION",
                                          "Easting", "Northing", "YEAR_WORKS", "AREA_HA", "LENGTH_KM", "Easting",
                                          "Northing", "BUFFER_TYPE"])
    arcpy.edit.Densify("DAP_buff100", "ANGLE", "", "", "0.10")
    arcpy.analysis.Buffer("DAP_temp", "DAP_buff250", "250 meters", dissolve_option="LIST",
                          dissolve_field=["DAP_REF_NO", "DAP_NAME", "DISTRICT", "SCHEDULE", "RISK_LVL", "DESCRIPTION",
                                          "Easting", "Northing", "YEAR_WORKS", "AREA_HA", "LENGTH_KM", "Easting",
                                          "Northing", "BUFFER_TYPE"])
    arcpy.edit.Densify("DAP_buff250", "ANGLE", "", "", "0.10")
    arcpy.analysis.Buffer("DAP_temp", "DAP_buff500", "500 meters", dissolve_option="LIST",
                          dissolve_field=["DAP_REF_NO", "DAP_NAME", "DISTRICT", "SCHEDULE", "RISK_LVL", "DESCRIPTION",
                                          "Easting", "Northing", "YEAR_WORKS", "AREA_HA", "LENGTH_KM", "Easting",
                                          "Northing", "BUFFER_TYPE"])
    arcpy.edit.Densify("DAP_buff500", "ANGLE", "", "", "0.10")
    if mode == 'JFMP':
        # CH JFMP buffers
        arcpy.analysis.Buffer("DAP_temp", "DAP_OSbuff500", "500 meters",
                              line_side="OUTSIDE_ONLY")  # 500m buffer outside of the Burn included only - for contingency area
        arcpy.edit.Densify("DAP_OSbuff500", "ANGLE", "", "", "0.10")
        arcpy.analysis.Buffer("DAP_buff500", "DAP_OSbuff1000", "500 meters",
                              line_side="OUTSIDE_ONLY")  # 500m buffer on top of the Burn's 500m buffer - for contingency area buffer
        arcpy.edit.Densify("DAP_OSbuff1000", "ANGLE", "", "", "0.10")
        arcpy.management.CalculateField("DAP_temp", "BUFFER_TYPE", "'JFMP Area'", "PYTHON3")
        arcpy.management.CalculateField("DAP_OSbuff500", "BUFFER_TYPE",
                                        "'JFMP Contingency Area (500m from planned burn)'", "PYTHON3")
        arcpy.management.CalculateField("DAP_OSbuff1000", "BUFFER_TYPE",
                                        "'JFMP Contingency Buffer(500m from contingency area)'", "PYTHON3")
        arcpy.management.Merge(["DAP_temp", "DAP_OSbuff500", "DAP_OSbuff1000"],
                               "XJFMP_CHContingencyBuff")  # creates a buffer of JFMP + contingency + contingency buffer for CH checks
        arcpy.management.Dissolve("XJFMP_CHContingencyBuff", "JFMP_CHContingencyBuff",
                                  ["DAP_REF_NO", "DAP_NAME", "DISTRICT", "SCHEDULE", "RISK_LVL", "DESCRIPTION",
                                   "Easting", "Northing", "YEAR_WORKS", "AREA_HA", "LENGTH_KM", "BUFFER_TYPE"])

        # Biodiversity JFMP Buffers
        arcpy.management.CalculateField("DAP_buff100", "BUFFER_TYPE", "'JFMP Area'", "PYTHON3")
        arcpy.analysis.Buffer("DAP_buff100", "DAP_OSbuff900", "900 meters",
                              line_side="OUTSIDE_ONLY")  # Biodiversity contingency buffer: 900m from 100m buffer (1000m total)
        arcpy.edit.Densify("DAP_OSbuff900", "ANGLE", "", "", "0.10")
        arcpy.management.CalculateField("DAP_OSbuff900", "BUFFER_TYPE", "'JFMP Contingency Area'", "PYTHON3")
        arcpy.management.Merge(["DAP_buff100", "DAP_OSbuff900"],
                               "XJFMP_BioContingencyBuff")  # creates a shape for JFMP biodiversity checks
        arcpy.management.Dissolve("XJFMP_BioContingencyBuff", "JFMP_BioContingencyBuff",
                                  ["DAP_REF_NO", "DAP_NAME", "DISTRICT", "SCHEDULE", "RISK_LVL", "DESCRIPTION",
                                   "Easting", "Northing", "YEAR_WORKS", "AREA_HA", "LENGTH_KM", "BUFFER_TYPE"])

    # Overlay DAP shape with input datasets...
    for theme in theme_list:

        # --------------------------------------------------------------------
        # SUMMARY   -> intersecting all relevant input layers for summary theme

        if theme == "summary":
            if RiskLevel != "LRLI":  # only run this theme for DAP activities where risk is not LRLI
                print("")
                print("  --> creating combined layer for SUMMARY table...")

                # NOTE: for summary info, the original DAP boundary is used instead of the buffered boundary

                # compiling a single layer from derived PLM overlays
                print("   Creating PLM overlay data")
                try:
                    arcpy.analysis.SpatialJoin("DAP_buff1", in_PLMOVRLAY, "tempplmovrlay",
                                               join_operation="JOIN_ONE_TO_MANY",
                                               join_type="KEEP_ALL", match_option="INTERSECT")
                    arcpy.management.AlterField("tempplmovrlay", "LABEL", "PLM_OVERLAY", "PLM_OVERLAY")
                    arcpy.management.Dissolve("tempplmovrlay", "tempplm100",
                                              ["DAP_REF_NO", "DAP_NAME", "DISTRICT", "DESCRIPTION", "RISK_LVL",
                                               "SCHEDULE", "Easting", "Northing", "AREA_HA", "LENGTH_KM",
                                               "PLM_OVERLAY"])  # "WM_NAME"
                    arcpy.management.Delete("tempplmovrlay")

                except:
                    print("Error in processing.  Running except clause")
                    if arcpy.Exists("tempplm100"):
                        arcpy.management.Delete("tempplm100")
                    arcpy.management.Copy("DAP_buff1", "tempplm100")
                    arcpy.management.AddField("tempplm100", "PLM_OVERLAY", "TEXT", "15")
                    arcpy.management.CalculateField("templyr", "PLM_OVERLAY", '"OVERLAY ERROR"', "PYTHON3")
                    arcpy.management.Delete("templyr")

                # Combining Forest type layer with state background
                print("   Creating forestype data")
                try:
                    arcpy.analysis.SpatialJoin("DAP_buff1", in_FTYPE, "tempftype", join_operation="JOIN_ONE_TO_MANY",
                                               join_type="KEEP_ALL", match_option="INTERSECT")
                    arcpy.management.Dissolve("tempftype", "tempftyped",
                                              ["DAP_REF_NO", "DAP_NAME", "DISTRICT", "DESCRIPTION", "RISK_LVL",
                                               "SCHEDULE", "Easting", "Northing", "AREA_HA", "LENGTH_KM", "X_DESC"])

                    fclist = "xtempftype", "tempftype"
                    for fc in fclist:
                        arcpy.management.Delete(fc)
                    arcpy.management.Rename("tempftyped", "tempftype")
                except:
                    print("Error in processing.  Running except clause")
                    if arcpy.Exists("tempftype"):
                        arcpy.management.Delete("tempftype")
                    arcpy.management.Copy("DAP_buff1", "tempftype")
                    arcpy.management.AddField("tempftype", "X_DESC", "TEXT" "100")
                    arcpy.management.MakeFeatureLayer("tempftype", "templyr")
                    arcpy.management.CalculateField("templyr", "X_DESC", '"OVERLAY ERROR"', "PYTHON3")
                    arcpy.management.Delete("templyr")

                # Compiling usable PLM25 layer with state background
                print("   Creating PLM25 data")
                try:
                    arcpy.analysis.SpatialJoin("DAP_buff1", in_PLM, "tempplm", join_operation="JOIN_ONE_TO_MANY",
                                               join_type="KEEP_ALL", match_option="INTERSECT")
                    arcpy.management.Dissolve("tempplm", "tempplmd",
                                              ["DAP_REF_NO", "DAP_NAME", "DISTRICT", "DESCRIPTION", "RISK_LVL",
                                               "SCHEDULE", "Easting", "Northing", "AREA_HA", "LENGTH_KM", "MMTGEN",
                                               "MNG_SPEC", "ACT"])

                    fclist = "xtempplm", "tempplm"
                    for fc in fclist:
                        arcpy.management.Delete(fc)
                    arcpy.management.Rename("tempplmd", "tempplm")

                except:
                    print("Error in processing.  Running except clause")
                    if arcpy.Exists("tempplm"):
                        arcpy.management.Delete("tempplm")
                    arcpy.management.Copy("DAP_buff1", "tempplm")
                    arcpy.management.AddField("tempplm", "MMTGEN", "TEXT", "15")
                    arcpy.management.AddField("tempplm", "MNG_SPEC", "TEXT", "15")
                    arcpy.management.MakeFeatureLayer("tempplm", "templyr")
                    arcpy.management.CalculateField("templyr", "MMTGEN", '"OVERLAY ERROR"', "PYTHON3")
                    arcpy.management.CalculateField("templyr", "MNG_SPEC", '"OVERLAY ERROR"', "PYTHON3")
                    arcpy.management.Delete("templyr")

                # Compiling usable CROWN PARCEL layer with state background
                print("   Creating Crown parcels data")
                try:
                    arcpy.analysis.SpatialJoin("DAP_buff1", in_PARCEL, "temppcl", join_operation="JOIN_ONE_TO_MANY",
                                               join_type="KEEP_ALL", match_option="INTERSECT")
                    arcpy.management.Dissolve("temppcl", "temppcld",
                                              ["DAP_REF_NO", "DAP_NAME", "DISTRICT", "DESCRIPTION", "RISK_LVL",
                                               "SCHEDULE", "Easting", "Northing", "AREA_HA", "LENGTH_KM", "ALLOTMENT",
                                               "SEC", "PARISH_CODE", "P_NUMBER", "TOWNSHIP_CODE"])

                    fclist = "xtemppcl", "temppcl"
                    for fc in fclist:
                        arcpy.management.Delete(fc)
                    arcpy.management.Rename("temppcld", "temppcl")

                except:
                    print("Error in processing.  Running except clause")
                    if arcpy.Exists("temppcl"):
                        arcpy.management.Delete("temppcl")
                    arcpy.management.Copy("DAP_buff1", "temppcl")
                    arcpy.management.AddField("temppcl", "ALLOTMENT", "TEXT", "15")
                    arcpy.management.AddField("temppcl", "SEC", "TEXT", "15")
                    arcpy.management.AddField("temppcl", "PARISH_CODE", "TEXT", "15")
                    arcpy.management.AddField("temppcl", "P_NUMBER", "TEXT", "15")
                    arcpy.management.AddField("temppcl", "TOWNSHIP_CODE", "TEXT", "15")
                    arcpy.management.MakeFeatureLayer("temppcl", "templyr")
                    arcpy.management.CalculateField("templyr", "ALLOTMENT", '"OVERLAY ERROR"', "PYTHON3")
                    arcpy.management.CalculateField("templyr", "SEC", '"OVERLAY ERROR"', "PYTHON3")
                    arcpy.management.CalculateField("templyr", "PARISH_CODE", '"OVERLAY ERROR"', "PYTHON3")
                    arcpy.management.CalculateField("templyr", "P_NUMBER", '"OVERLAY ERROR"', "PYTHON3")
                    arcpy.management.CalculateField("templyr", "TOWNSHIP_CODE", '"OVERLAY ERROR"', "PYTHON3")
                    arcpy.management.Delete("templyr")

                # Creating clipped Planning zone layer
                print("   Processing Planning Zones data...")
                arcpy.analysis.SpatialJoin("DAP_buff1", in_PLAN, "tempplan", join_operation="JOIN_ONE_TO_MANY",
                                           join_type="KEEP_ALL", match_option="INTERSECT")
                arcpy.management.Dissolve("tempplan", "temppland",
                                          ["DAP_REF_NO", "DAP_NAME", "DISTRICT", "DESCRIPTION", "RISK_LVL", "SCHEDULE",
                                           "Easting", "Northing", "AREA_HA", "LENGTH_KM", "ZONE_CODE", "LGA"], "",
                                          "MULTI_PART")

                arcpy.management.Delete("tempplan")
                arcpy.management.Rename("temppland", "tempplan")

                # Creating clipped Planning overlay layer
                print("   Processing Planning Overlay data...")
                arcpy.analysis.SpatialJoin("DAP_buff1", in_PLOLAY, "tempolay", join_operation="JOIN_ONE_TO_MANY",
                                           join_type="KEEP_ALL", match_option="INTERSECT")
                arcpy.management.Dissolve("tempolay", "tempolayd",
                                          ["DAP_REF_NO", "DAP_NAME", "DISTRICT", "DESCRIPTION", "RISK_LVL", "SCHEDULE",
                                           "Easting", "Northing", "AREA_HA", "LENGTH_KM", "ZONE_CODE"], "",
                                          "MULTI_PART")
                arcpy.management.AlterField("tempolayd", "ZONE_CODE", "OVERLAY", "OVERLAY")

                arcpy.management.Delete("tempolay")
                arcpy.management.Rename("tempolayd", "tempolay")

                # Processing Native Title Layer
                print("   Processing Native Title data...")
                arcpy.analysis.SpatialJoin("DAP_buff1", in_NT, "xtempnt", join_operation="JOIN_ONE_TO_MANY",
                                           join_type="KEEP_ALL", match_option="INTERSECT")
                arcpy.management.Dissolve("xtempnt", "tempnt",
                                          ["DAP_REF_NO", "DAP_NAME", "DISTRICT", "DESCRIPTION", "RISK_LVL", "SCHEDULE",
                                           "Easting", "Northing", "AREA_HA", "LENGTH_KM", "NT_Name", "Tribunal_No",
                                           "NT_STATUS"], "", "MULTI_PART")
                # arcpy.management.AlterField("tempnt", "Tribunal_No", "NT_Tribunal_No","NT_Tribunal_No")

                arcpy.management.Delete("xtempnt")

                # processing all summary data...
                print("   Combining all summary inputs")

                # This section below conducts a search cursor to find unique values for each activity ID, then directly imports them into the Summary Table

                TabFields = ["DAP_REF_NO", "DAP_NAME", "DESCRIPTION", "RISK_LVL", "SCHEDULE", "AREA_HA", "LENGTH_KM",
                             "Easting", "Northing", "DISTRICT", "SOIL_DISTURB", "NATIVE_VEG", "WORKS_TYPE_1",
                             "WORKS_TYPE_2", "WORKS_TYPE_3"]  ## "EASTING", "NORTHING", #WORKS_TYPE?
                with arcpy.da.SearchCursor("DAP_buff1", TabFields, "RISK_LVL <> 'LRLI'") as sCur:
                    with arcpy.da.InsertCursor(sumtab, ["UNIQUE_ID", "DISTRICT", "SITE_NAME", "SCHEDULE", "DESCRIPTION",
                                                        "STATUS", "AREA_HA", "LENGTH_KM", "Easting", "Northing",
                                                        "SOIL_DISTURB", "NATIVE_VEG", "WORKS_TYPE_1", "WORKS_TYPE_2",
                                                        "WORKS_TYPE_3"]) as iCur:  # "EASTING", "NORTHING",
                        for rows in sCur:
                            # Use insert cursor to add and populate table records - match the row number from DAP_tab in same sequence as the fields for the Summary Table
                            iCur.insertRow((rows[0], rows[9], rows[1], rows[4], rows[2], rows[3], rows[5], rows[6],
                                            rows[7], rows[8], rows[10], rows[11], rows[12], rows[13], rows[14]))

                pcl = "temppcl"
                olay = "tempolay"
                plan = "tempplan"
                plm = "tempplm"
                plm100 = "tempplm100"
                ftype = "tempftype"
                nt = "tempnt"

                IDs = [row[0] for row in arcpy.da.SearchCursor("DAP_buff1", reference_field, "RISK_LVL <> 'LRLI'")]
                UniqueID = set(IDs)

                for ID in UniqueID:
                    exp = """{0} = '{1}'""".format(arcpy.AddFieldDelimiters(sumtab, reference_field), ID)
                    print("Processing values for " + ID)

                    ##'''
                    ##public crown land data
                    ##'''
                    # Allotment
                    Allot = [row[0] for row in arcpy.da.SearchCursor(pcl, "ALLOTMENT", exp) if
                             row not in (None, '', 'None', ' ')]  # expression used for searchcusor is the issue
                    if Allot:  # checks if it's not an empty list, the loop will be bypassed if the list is empty
                        UAllot = set(Allot)  # get unique value set
                        UAllot.discard(None)
                        UAllot.discard("None")
                        UAllot.discard("")
                        UAllot.discard(" ")
                        ValueAllot = '; '.join([str(item) for item in
                                                UAllot])  # convert into a single string for importing into Summary Table

                        with arcpy.da.UpdateCursor(sumtab, "C_ALLOT", "UNIQUE_ID = " + "'" + ID + "'") as cursor:
                            for row in cursor:
                                row[0] = str(ValueAllot)
                                cursor.updateRow(row)
                        del UAllot
                        del ValueAllot

                    del Allot

                    # SEC
                    SEC = [row[0] for row in arcpy.da.SearchCursor(pcl, "SEC", exp) if
                           (row is not None or row not in ('', 'None', ' '))]
                    if SEC:
                        USEC = set(SEC)  # get unique value set
                        USEC.discard(None)
                        USEC.discard("None")
                        USEC.discard("")
                        USEC.discard(" ")
                        ValueSEC = '; '.join([str(item) for item in
                                              USEC])  # convert into a single string for importing into Summary Table

                        with arcpy.da.UpdateCursor(sumtab, "C_SEC", "UNIQUE_ID = " + "'" + ID + "'") as cursor:
                            for row in cursor:
                                row[0] = str(ValueSEC)
                                cursor.updateRow(row)
                        del USEC
                        del ValueSEC

                    del SEC

                    # PARISH
                    Parish = [row[0] for row in arcpy.da.SearchCursor(pcl, "PARISH_CODE", exp) if
                              row not in (None, '', 'None', ' ')]
                    if Parish:
                        UParish = set(Parish)  # get unique value set
                        UParish.discard(None)
                        UParish.discard("None")
                        UParish.discard("")
                        UParish.discard(" ")
                        ValueParish = '; '.join([str(item) for item in
                                                 UParish])  # convert into a single string for importing into Summary Table

                        with arcpy.da.UpdateCursor(sumtab, "C_PARISH", "UNIQUE_ID = " + "'" + ID + "'") as cursor:
                            for row in cursor:
                                row[0] = str(ValueParish)
                                cursor.updateRow(row)
                        del UParish
                        del ValueParish
                    del Parish

                    # P_Number
                    PNum = [row[0] for row in arcpy.da.SearchCursor(pcl, "P_NUMBER", exp) if
                            row not in (None, '', 'None', ' ')]
                    if PNum:
                        UPNum = set(PNum)  # get unique value set
                        UPNum.discard(None)
                        UPNum.discard("None")
                        UPNum.discard("")
                        UPNum.discard(" ")
                        ValuePNum = '; '.join([str(item) for item in
                                               UPNum])  # convert into a single string for importing into Summary Table
                        try:
                            with arcpy.da.UpdateCursor(sumtab, "C_PNUM", "UNIQUE_ID = " + "'" + ID + "'") as cursor:
                                for row in cursor:
                                    row[0] = str(ValuePNum)
                                    cursor.updateRow(row)
                        except:
                            with arcpy.da.UpdateCursor(sumtab, "C_PNUM", "UNIQUE_ID = " + "'" + ID + "'") as cursor:
                                for row in cursor:
                                    row[0] = "Area too large to list all relevant items, check spatial layer"
                                    cursor.updateRow(row)
                        del UPNum
                        del ValuePNum
                    del PNum

                    # Township
                    Tship = [row[0] for row in arcpy.da.SearchCursor(pcl, "TOWNSHIP_CODE", exp) if
                             row not in (None, '', 'None', ' ')]
                    if Tship:
                        UTship = set(Tship)  # get unique value set
                        UTship.discard(None)
                        UTship.discard("None")
                        UTship.discard("")
                        UTship.discard(" ")
                        ValueTship = '; '.join([str(item) for item in
                                                UTship])  # convert into a single string for importing into Summary Table

                        with arcpy.da.UpdateCursor(sumtab, "C_TSHIP", "UNIQUE_ID = " + "'" + ID + "'") as cursor:
                            for row in cursor:
                                row[0] = str(ValueTship)
                                cursor.updateRow(row)
                        del UTship
                        del ValueTship
                    del Tship

                    ##'''
                    ##public land overlay data
                    ##'''
                    # Overlay
                    Olay = [row[0] for row in arcpy.da.SearchCursor(olay, "OVERLAY", exp) if
                            row not in (None, '', 'None', ' ')]
                    if Olay:
                        UOlay = set(Olay)  # get unique value set
                        UOlay.discard(None)
                        UOlay.discard("None")
                        UOlay.discard("")
                        UOlay.discard(" ")
                        ValueOlay = '; '.join([str(item) for item in
                                               UOlay])  # convert into a single string for importing into Summary Table

                        with arcpy.da.UpdateCursor(sumtab, "OVERLAY", "UNIQUE_ID = " + "'" + ID + "'") as cursor:
                            for row in cursor:
                                row[0] = str(ValueOlay)
                                cursor.updateRow(row)
                        del UOlay
                        del ValueOlay
                    del Olay

                    ##'''
                    ##planning data
                    ##'''
                    # Plan Zone
                    Plan = [row[0] for row in arcpy.da.SearchCursor(plan, "ZONE_CODE", exp) if
                            row not in (None, '', 'None', ' ')]
                    if Plan:
                        UPlan = set(Plan)  # get unique value set
                        UPlan.discard(None)
                        UPlan.discard("None")
                        UPlan.discard("")
                        UPlan.discard(" ")
                        ValuePlan = '; '.join([str(item) for item in
                                               UPlan])  # convert into a single string for importing into Summary Table

                        with arcpy.da.UpdateCursor(sumtab, "PL_ZONE", "UNIQUE_ID = " + "'" + ID + "'") as cursor:
                            for row in cursor:
                                row[0] = str(ValuePlan)
                                cursor.updateRow(row)
                        del UPlan
                        del ValuePlan
                    del Plan

                    # LGA
                    LGA = [row[0] for row in arcpy.da.SearchCursor(plan, "LGA", exp) if
                           row not in (None, '', 'None', ' ')]
                    if LGA:
                        ULGA = set(LGA)  # get unique value set
                        ULGA.discard(None)
                        ULGA.discard("None")
                        ULGA.discard("")
                        ULGA.discard(" ")
                        ValueLGA = '; '.join([str(item) for item in
                                              ULGA])  # convert into a single string for importing into Summary Table

                        with arcpy.da.UpdateCursor(sumtab, "LGA", "UNIQUE_ID = " + "'" + ID + "'") as cursor:
                            for row in cursor:
                                row[0] = str(ValueLGA)
                                cursor.updateRow(row)
                        del ULGA
                        del ValueLGA
                    del LGA

                    ##'''
                    ##public land mgt 25 data
                    ##'''
                    # Land Status
                    Land = [row[0] for row in arcpy.da.SearchCursor(plm, "MMTGEN", exp) if
                            row not in (None, '', 'None', ' ')]
                    if Land:
                        ULand = set(Land)  # get unique value set
                        ULand.discard(None)
                        ULand.discard("None")
                        ULand.discard("")
                        ULand.discard(" ")
                        ValueLand = '; '.join([str(item) for item in
                                               ULand])  # convert into a single string for importing into Summary Table

                        with arcpy.da.UpdateCursor(sumtab, "LAND_STATUS", "UNIQUE_ID = " + "'" + ID + "'") as cursor:
                            for row in cursor:
                                row[0] = str(ValueLand)
                                cursor.updateRow(row)
                        del ULand
                        del ValueLand
                    del Land

                    # Land Manager
                    Mngr = [row[0] for row in arcpy.da.SearchCursor(plm, "MNG_SPEC", exp) if
                            row not in (None, '', 'None', ' ')]
                    if Mngr:
                        UMngr = set(Mngr)  # get unique value set
                        UMngr.discard(None)
                        UMngr.discard("None")
                        UMngr.discard("")
                        UMngr.discard(" ")
                        ValueMngr = '; '.join([str(item) for item in
                                               UMngr])  # convert into a single string for importing into Summary Table

                        with arcpy.da.UpdateCursor(sumtab, "LAND_MANGR", "UNIQUE_ID = " + "'" + ID + "'") as cursor:
                            for row in cursor:
                                row[0] = str(ValueMngr)
                                cursor.updateRow(row)
                        del UMngr
                        del ValueMngr
                    del Mngr

                    # Act
                    Act = [row[0] for row in arcpy.da.SearchCursor(plm, "ACT", exp) if
                           row not in (None, '', 'None', ' ')]
                    if Act:
                        UAct = set(Act)  # get unique value set
                        UAct.discard(None)
                        UAct.discard("None")
                        UAct.discard("")
                        UAct.discard(" ")
                        ValueAct = '; '.join([str(item) for item in
                                              UAct])  # convert into a single string for importing into Summary Table

                        with arcpy.da.UpdateCursor(sumtab, "ACT", "UNIQUE_ID = " + "'" + ID + "'") as cursor:
                            for row in cursor:
                                row[0] = str(ValueAct)
                                cursor.updateRow(row)
                        del UAct
                        del ValueAct
                    del Act

                    ##'''
                    ##public land overlay 100 data
                    ##'''
                    # plm overlay
                    PLMO = [row[0] for row in arcpy.da.SearchCursor(plm100, "PLM_OVERLAY", exp) if
                            row not in (None, '', 'None', ' ')]
                    if PLMO:
                        UPLMO = set(PLMO)  # get unique value set
                        UPLMO.discard(None)
                        UPLMO.discard("None")
                        UPLMO.discard("")
                        UPLMO.discard(" ")
                        ValuePLMO = '; '.join([str(item) for item in
                                               UPLMO])  # convert into a single string for importing into Summary Table

                        with arcpy.da.UpdateCursor(sumtab, "PLM_OVERLAY", "UNIQUE_ID = " + "'" + ID + "'") as cursor:
                            for row in cursor:
                                row[0] = str(ValuePLMO)
                                cursor.updateRow(row)
                        del UPLMO
                        del ValuePLMO
                    del PLMO

                    ##'''
                    ##forest type data
                    ##'''
                    # forest type
                    Frst = [row[0] for row in arcpy.da.SearchCursor(ftype, "X_DESC", exp) if
                            row not in (None, '', 'None', ' ')]
                    if Frst:
                        UFrst = set(Frst)  # get unique value set
                        UFrst.discard(None)
                        UFrst.discard("None")
                        UFrst.discard("")
                        UFrst.discard(" ")
                        ValueFrst = '; '.join([str(item) for item in
                                               UFrst])  # convert into a single string for importing into Summary Table

                        with arcpy.da.UpdateCursor(sumtab, "FOR_TYPE", "UNIQUE_ID = " + "'" + ID + "'") as cursor:
                            for row in cursor:
                                row[0] = str(ValueFrst)
                                cursor.updateRow(row)
                        del UFrst
                        del ValueFrst
                    del Frst

                    ##'''
                    ##Native Title
                    ##'''
                    # NT Tribunal Number

                    NTno = [row[0] for row in arcpy.da.SearchCursor(nt, "Tribunal_No", exp) if
                            row not in (None, '', 'None', ' ')]
                    if NTno:
                        UNTno = set(NTno)  # get unique value set
                        UNTno.discard(None)
                        UNTno.discard("None")
                        UNTno.discard("")
                        UNTno.discard(" ")
                        ValueNTno = '; '.join([str(item) for item in
                                               UNTno])  # convert into a single string for importing into Summary Table

                        with arcpy.da.UpdateCursor(sumtab, "NT_Tribunal_No", "UNIQUE_ID = " + "'" + ID + "'") as cursor:
                            for row in cursor:
                                row[0] = str(ValueNTno)
                                cursor.updateRow(row)
                        del UNTno
                        del ValueNTno
                    del NTno

                    # NT Name
                    NTname = [row[0] for row in arcpy.da.SearchCursor(nt, "NT_Name", exp) if
                              row not in (None, '', 'None', ' ')]
                    if NTname:
                        UNTname = set(NTname)  # get unique value set
                        UNTname.discard(None)
                        UNTname.discard("None")
                        UNTname.discard("")
                        UNTname.discard(" ")
                        ValueNTname = '; '.join([str(item) for item in
                                                 UNTname])  # convert into a single string for importing into Summary Table

                        with arcpy.da.UpdateCursor(sumtab, "NT_Name", "UNIQUE_ID = " + "'" + ID + "'") as cursor:
                            for row in cursor:
                                row[0] = str(ValueNTname)
                                cursor.updateRow(row)
                        del UNTname
                        del ValueNTname
                    del NTname

                    # NT Status
                    NTstatus = [row[0] for row in arcpy.da.SearchCursor(nt, "NT_STATUS", exp) if
                                row not in (None, '', 'None', ' ')]
                    if NTstatus:
                        UNTstat = set(NTstatus)  # get unique value set
                        UNTstat.discard(None)
                        UNTstat.discard("None")
                        UNTstat.discard("")
                        UNTstat.discard(" ")
                        ValueNTstat = '; '.join([str(item) for item in
                                                 UNTstat])  # convert into a single string for importing into Summary Table

                        with arcpy.da.UpdateCursor(sumtab, "NT_STATUS", "UNIQUE_ID = " + "'" + ID + "'") as cursor:
                            for row in cursor:
                                row[0] = str(ValueNTstat)
                                cursor.updateRow(row)
                        del UNTstat
                        del ValueNTstat
                    del NTstatus

                # Blank values to be filled in with 'NA'
                with arcpy.da.UpdateCursor(sumtab, ["LAND_STATUS", "LAND_MANGR", "FOR_TYPE", "PLM_OVERLAY", "OVERLAY",
                                                    "PL_ZONE",
                                                    "LGA", "C_ALLOT", "C_SEC", "C_PARISH", "C_PNUM", "C_TSHIP", "ACT",
                                                    "NT_Tribunal_No", "NT_Name", "NT_STATUS"]) as cursor:
                    for row in cursor:
                        if (row[0] == '' or row[0] is None):
                            row[0] = 'NA'
                        if (row[1] == '' or row[1] is None):
                            row[1] = 'NA'
                        if (row[2] == '' or row[2] is None):
                            row[2] = 'NA'
                        if (row[3] == '' or row[3] is None):
                            row[3] = 'NA'
                        if (row[4] == '' or row[4] is None):
                            row[4] = 'NA'
                        if (row[5] == '' or row[5] is None):
                            row[5] = 'NA'
                        if (row[6] == '' or row[6] is None):
                            row[6] = 'NA'
                        if (row[7] == '' or row[7] is None):
                            row[7] = 'NA'
                        if (row[8] == '' or row[8] is None):
                            row[8] = 'NA'
                        if (row[9] == '' or row[9] is None):
                            row[9] = 'NA'
                        if (row[10] == '' or row[10] is None):
                            row[10] = 'NA'
                        if (row[11] == '' or row[11] is None):
                            row[11] = 'NA'
                        if (row[12] == '' or row[12] is None):
                            row[12] = 'NA'
                        if (row[13] == '' or row[13] is None):
                            row[13] = 'NA'
                        if (row[14] == '' or row[14] is None):
                            row[14] = 'NA'
                        if (row[15] == '' or row[15] is None):
                            row[15] = 'NA'
                        cursor.updateRow(row)

                # cleaning up...
                print("Cleaning up temp Summary data...")
                for fc in "DAP_theme", "DAP_single", "DAP_layer", "DAP_themecl", "tempplm", "tempftype", "tempplm100", "tempplan", "tempolay", "temppcl", "tempnt":
                    if arcpy.Exists(fc):
                        arcpy.management.Delete(fc)


        # ---------------------------------------------------------------------
        # FORESTS   -> intersecting all relevant input layers for forests theme

        elif theme == "forests":
            print("")
            print("  --> creating combined layer for FORESTS table")
            # Creating a series of temporary layers converting points to polys and combining them
            # These layers will be deleted at the end of the script

            if RiskLevel != "LRLI":  # If Risk level is not LRLI, run the following values checks (FOR DAP/NBFT only)
                # Creating Forest Management Zones layer from CSDL source data
                print("   Processing FMZ data...")
                arcpy.analysis.SpatialJoin("DAP_buff1", in_FMZ, "xtempfmz", join_operation="JOIN_ONE_TO_MANY",
                                           join_type="KEEP_ALL", match_option="INTERSECT")
                arcpy.management.AddField("xtempfmz", "Value_ID", "TEXT", "", "", 100)
                arcpy.management.MakeFeatureLayer("xtempfmz", "tempfmzlyr")
                arcpy.management.SelectLayerByAttribute("tempfmzlyr", "NEW_SELECTION",
                                                        "FMZ IS NOT NULL AND DETAILNO IS NOT NULL")
                arcpy.management.CalculateField("xtempfmz", "Value_ID", "!FMZ! + !DETAILNO!", "PYTHON3")
                arcpy.management.SelectLayerByAttribute("tempfmzlyr", "NEW_SELECTION",
                                                        "FMZ IS NOT NULL AND FMZ_NO IS NOT NULL AND DETAILNO IS NULL")
                arcpy.management.CalculateField("xtempfmz", "Value_ID", "!FMZ! + !FMZ_NO!", "PYTHON3")
                arcpy.management.SelectLayerByAttribute("tempfmzlyr", "NEW_SELECTION", "FMZDIS IN('SPZ', 'SMZ')")
                arcpy.management.Dissolve("tempfmzlyr", "tempfmz",
                                          ["DAP_REF_NO", "DAP_NAME", "DISTRICT", "DESCRIPTION", "RISK_LVL", "FMZDIS",
                                           "DESC1", "Value_ID"], "", "MULTI_PART")

                ###### modify & simplify FMZ attributes
                if int(arcpy.management.GetCount("tempfmzlyr")[0]) > 0:
                    arcpy.management.AddField("tempfmz", "Value_Type", "TEXT", "", "", 50)
                    arcpy.management.CalculateField("tempfmz", "Value_Type", "'FMZ'", "PYTHON3")
                    arcpy.management.AlterField("tempfmz", "FMZDIS", "Value")
                    arcpy.management.AlterField("tempfmz", "DESC1", "Value_Description")

                for fc in "xtempfmz", "tempfmzlyr":
                    arcpy.management.Delete(fc)

                ##                  # Creating clipped Old Growth layer
                ##                    print "   Processing Old Growth data..."
                ##                    arcpy.analysis.SpatialJoin("DAP_buff1", in_OLDGROWTH, "xtempog", join_operation="JOIN_ONE_TO_MANY",
                ##                           join_type="KEEP_ALL", match_option="INTERSECT")
                ##                    arcpy.management.AddField("xtempog", "Value_Type", "TEXT", "", "", 50)
                ##                    arcpy.management.CalculateField("xtempog", "Value_Type", '"Modelled Old Growth"', "PYTHON3")
                ##                    arcpy.management.AlterField("xtempog", "X_OGDESC", "Value")
                ##                    arcpy.management.Dissolve("xtempog", "tempog", ["DAP_REF_NO","DAP_NAME","DESCRIPTION","Value","Value_Type"], "", "MULTI_PART")
                ##
                ##                    arcpy.management.Delete("xtempog")

                # Creating statewide Huts layer - convert points to polys to work in Intersect and change a field name
                if not arcpy.Exists("temphutsv"):
                    print("   Creating statewide Huts data...")
                    arcpy.analysis.Buffer(in_HUTS, "temphutsv", "10 meter", "FULL", "ROUND", "LIST",
                                          ["NAME", "EASTING", "NORTHING"])
                    arcpy.management.AlterField("temphutsv", "NAME", "Value", "Value")
                    arcpy.management.AddField("temphutsv", "Value_Type", "TEXT", "", "", 50)
                    arcpy.management.CalculateField("temphutsv", "Value_Type", '"Alpine Hut"', "PYTHON3")
                    arcpy.management.AlterField("temphutsv", "EASTING", "X", "X")
                    arcpy.management.AlterField("temphutsv", "NORTHING", "Y", "Y")

                print("   Processing huts data...")
                arcpy.analysis.SpatialJoin("DAP_buff1", "temphutsv", "xtemphuts", join_operation="JOIN_ONE_TO_MANY",
                                           join_type="KEEP_ALL", match_option="INTERSECT")
                arcpy.management.Dissolve("xtemphuts", "temphuts",
                                          ["DAP_REF_NO", "DAP_NAME", "DISTRICT", "DESCRIPTION", "RISK_LVL", "Value",
                                           "Value_Type", "X", "Y"], "", "MULTI_PART")
                arcpy.management.Delete("xtemphuts")

                # Creating Giant tree layer - convert points to polys to work in Intersect and change a field name, dissolve on required field
                if not arcpy.Exists("tempgtreesv"):
                    print("   Creating statewide Giant Trees data...")
                    arcpy.analysis.Buffer(in_GTREES, "tempgtreesv", "10 meter", "FULL", "ROUND", "LIST", ["SOURCE"])
                    arcpy.management.AlterField("tempgtreesv", "SOURCE", "Value", "Value")
                    arcpy.management.CalculateField("tempgtreesv", "Value", '"Giant tree"', "PYTHON3")
                    arcpy.management.AddField("tempgtreesv", "Value_Type", "TEXT", "", "", 50)
                    arcpy.management.CalculateField("tempgtreesv", "Value_Type", '"Giant tree"', "PYTHON3")
                    arcpy.management.AddField("tempgtreesv", "X", "DOUBLE", "", "0", "10")
                    arcpy.management.AddField("tempgtreesv", "Y", "DOUBLE", "", "0", "10")
                    with arcpy.da.UpdateCursor("tempgtreesv", ['SHAPE@X', 'SHAPE@Y', "X", "Y"],
                                               spatial_reference=srfilez) as cursor:
                        for row in cursor:
                            row[2] = int(row[0])
                            row[3] = int(row[1])
                            cursor.updateRow(row)
                print("   Processing Giant trees data...")
                arcpy.analysis.SpatialJoin("DAP_buff1", "tempgtreesv", "xtempgtrees", join_operation="JOIN_ONE_TO_MANY",
                                           join_type="KEEP_ALL", match_option="INTERSECT")
                arcpy.management.Dissolve("xtempgtrees", "tempgtrees",
                                          ["DAP_REF_NO", "DAP_NAME", "DISTRICT", "DESCRIPTION", "RISK_LVL", "Value",
                                           "Value_Type", "X", "Y"], "", "MULTI_PART")
                arcpy.management.Delete("xtempgtrees")

                #  Creating Rainforest inputs
                ##                    print "   Processing Rainforest data..."
                ###                # Removing central Highlands areas from CSDL rainforest layer, will be using later reference_field data for this area.
                ##                    if not arcpy.Exists("RF_ALL"):
                ##                        print "   Creating one-off rainforest data.."
                ##                        arcpy.analysis.Erase(in_RFPOLY, in_RFCLIP, "xtemprfpoly")
                ##                        arcpy.management.Merge(["xtemprfpoly", in_RFPOLYCH], "RF_All") # in_RFMURR]
                ##                        arcpy.management.MakeFeatureLayer ("RF_All", "rflyr")
                ##                        arcpy.management.SelectLayerByAttribute ("rflyr", "NEW_SELECTION", "RF = 1")
                ##                        arcpy.management.CalculateField("rflyr", "EVC_RF", '"Cool Temperate RF"', "PYTHON3") #Central Highlands to be calculated as "cool temperature rainforest"
                ##                        #arcpy.management.SelectLayerByAttribute ("rflyr", "NEW_SELECTION", "ET_ID > 0")
                ##                        #arcpy.management.CalculateField("rflyr", "EVC_RF", '"Murrungowar RF"', "PYTHON3")
                ##                        arcpy.management.AlterField("RF_All", "EVC_RF", "Value", "Value")
                ##                        arcpy.management.AddField("RF_All", "Value_Type", "TEXT", 15)
                ##                        arcpy.management.CalculateField("RF_All", "Value_Type", '"Rainforest"', "PYTHON3")
                ##                        arcpy.management.SelectLayerByAttribute ("rflyr", "NEW_SELECTION", "RF = 0") #remove world polygons
                ##                        if int(arcpy.management.GetCount("rflyr")[0]) > 0:
                ##                            arcpy.management.DeleteFeatures("rflyr")
                ##
                ##                        for fc in "xtemprfpoly", "rflyr":
                ##                            arcpy.management.Delete(fc)
                ##
                ##                    arcpy.analysis.SpatialJoin("DAP_buff1", "RF_ALL", "xtemprf", join_operation="JOIN_ONE_TO_MANY",
                ##                           join_type="KEEP_ALL", match_option="INTERSECT")
                ##                    arcpy.management.Dissolve("xtemprf", "temprf", ["DAP_REF_NO","DAP_NAME","DISTRICT","DESCRIPTION","RISK_LVL","Value","Value_Type"], "", "MULTI_PART")
                ##                    arcpy.management.Delete("xtemprf")

                ##                     # Creating Rainforest sites of significance layer - add background poly, dissolve on required field
                ##                    if not arcpy.Exists("temprainss"):
                ##                        print "   Creating statewide Rainforest sites data..."
                ##                        arcpy.conversion.FeatureClassToFeatureClass(in_RFSOS, workpath + "\\" + outGDB, "temprainss", "\"SIGNIF\" = 'N'")
                ##                        arcpy.management.AlterField("temprainss", "SITE_NAME", "Value", "Value")
                ##                        arcpy.management.AddField("temprainss", "Value_Type", "TEXT", "", "", 50)
                ##                        arcpy.management.CalculateField("temprainss", "Value_Type", '"Rainforest Site of Significance"', "PYTHON3")
                ##
                ##                    print "   Processing Rainforest SOS data..."
                ##                    arcpy.analysis.SpatialJoin("DAP_buff1", "temprainss", "xtemprainfor", join_operation="JOIN_ONE_TO_MANY",
                ##                           join_type="KEEP_ALL", match_option="INTERSECT")
                ##                    arcpy.management.MakeFeatureLayer ("xtemprainfor", "rainlyr")
                ##                    arcpy.management.SelectLayerByAttribute ("rainlyr", "NEW_SELECTION", "Value NOT IN('', NULL)")
                ##                    arcpy.management.Dissolve("rainlyr", "temprainfor", ["DAP_REF_NO","DAP_NAME","DESCRIPTION","Value", "Value_Type"], "", "MULTI_PART")
                ##
                ##                    for fc in "xtemprainfor", "rainlyr":
                ##                        arcpy.management.Delete(fc)

                # Creating RecWeb layer - combine points, convert to polys for Intersect, dissolve on required field
                # NOTE recweb_sign must be the first in the list for merge, as it has the longest fields and values. This sets the schema for the others in the list
                if not arcpy.Exists("temprecwebv"):
                    print("   Creating statewide Recweb data...")
                    arcpy.management.Merge([in_RECWEBN, in_RECWEBA, in_RECWEBS, in_RECWEBH, in_RECWEBC], "recweball")
                    arcpy.analysis.Buffer("recweball", "temprecwebv", "10 meter", "FULL", "ROUND", "LIST",
                                          ["SERIAL_NO", "NAME", "COMMENTS", "FAC_TYPE", "ASSET_CLS"])
                    arcpy.management.AlterField("temprecwebv", "FAC_TYPE", "Value_Type", "Value_Type")
                    arcpy.management.AlterField("temprecwebv", "NAME", "Value", "Value")
                    arcpy.management.AddField("temprecwebv", "Value_ID", "TEXT", "", "", 100)
                    arcpy.management.CalculateField("temprecwebv", "Value_ID", "!SERIAL_NO!", "PYTHON3")
                    arcpy.management.AlterField("temprecwebv", "COMMENTS", "Value_Description",
                                                "Value_Description")  # Check if this is too long when merging? Otherwise place temprecweb first in the list. Could RECWEB_ASSET use ASSET_CLS + COMMENTS for its Value_Description?
                    arcpy.management.AddField("temprecwebv", "X", "DOUBLE", "", "0", "10")
                    arcpy.management.AddField("temprecwebv", "Y", "DOUBLE", "", "0", "10")
                    with arcpy.da.UpdateCursor("temprecwebv", ['SHAPE@X', 'SHAPE@Y', "X", "Y", "Value_Description"],
                                               spatial_reference=srfilez) as cursor:
                        for row in cursor:
                            row[2] = int(row[0])
                            row[3] = int(row[1])
                            if row[4] is None:
                                row[4] = " "
                            if len(row[4]) > 255:  # Shorten the long comment descriptions here
                                LongEntry = row[4]
                                row[4] = LongEntry[:255]
                            cursor.updateRow(row)
                    arcpy.management.Delete("recweball")

                print("   Processing Recweb data...")
                arcpy.analysis.SpatialJoin("DAP_buff1", "temprecwebv", "xtemprecweb", join_operation="JOIN_ONE_TO_MANY",
                                           join_type="KEEP_ALL", match_option="INTERSECT")
                arcpy.management.Dissolve("xtemprecweb", "temprecweb",
                                          ["DAP_REF_NO", "DAP_NAME", "DISTRICT", "DESCRIPTION", "RISK_LVL", "Value",
                                           "Value_Type", "Value_Description", "Value_ID", "X", "Y"], "", "MULTI_PART")
                arcpy.management.Delete("xtemprecweb")

                # Creating grazing license layer from CSDL source data
                print("   Processing Grazing License data...")
                if not arcpy.Exists("tempgraz"):
                    print("   Creating statewide grazing dataset...")
                    grazexp = "CLTEN_TENURE_CODE >= '100' and CLTEN_TENURE_CODE < '200'"
                    arcpy.Select_analysis(in_GLIC, "tempgraz", grazexp)
                    arcpy.management.JoinField("tempgraz", "CLTEN_TENURE_CODE", in_TCODE, "TENURE_CODE", "TENURE")
                    arcpy.management.AddField("tempgraz", "Value_Type", "TEXT", "", "", 50)
                    arcpy.management.CalculateField("tempgraz", "Value_Type", '"Grazing Licence"', "PYTHON3")
                    arcpy.management.AlterField("tempgraz", "CLTEN_TENURE_ID", "Value_Description", "Value_Description")
                    arcpy.management.AlterField("tempgraz", "TENURE", "Value", "Value")

                arcpy.analysis.SpatialJoin("DAP_buff1", "tempgraz", "xtempglic", join_operation="JOIN_ONE_TO_MANY",
                                           join_type="KEEP_ALL", match_option="INTERSECT")
                arcpy.management.Dissolve("xtempglic", "tempglic",
                                          ["DAP_REF_NO", "DAP_NAME", "DISTRICT", "DESCRIPTION", "RISK_LVL", "Value",
                                           "Value_Type", "Value_Description"], "", "MULTI_PART")
                arcpy.management.Delete("xtempglic")

            # Creating Pest Plant and Animal layer - convert points to polys to work in Intersect and change a field name, dissolve on required field
            print("   Processing Pests data...")

            arcpy.analysis.SpatialJoin("DAP_buff1", "temppestv", "xtemppest", join_operation="JOIN_ONE_TO_MANY",
                                       join_type="KEEP_ALL", match_option="INTERSECT")
            arcpy.management.Dissolve("xtemppest", "temppest",
                                      ["DAP_REF_NO", "DAP_NAME", "DISTRICT", "DESCRIPTION", "RISK_LVL", "Value",
                                       "Value_Type", "X", "Y"], "", "MULTI_PART")
            arcpy.management.Delete("xtemppest")

            # Creating Monitoring sites layer - add background poly, dissolve on required field
            print("   Processing monitoring sites data...")

            arcpy.analysis.SpatialJoin("DAP_buff1", "monforestfire", "xtempmon", join_operation="JOIN_ONE_TO_MANY",
                                       join_type="KEEP_ALL", match_option="INTERSECT")
            arcpy.management.Dissolve("xtempmon", "tempmon",
                                      ["DAP_REF_NO", "DAP_NAME", "DISTRICT", "DESCRIPTION", "RISK_LVL", "Value",
                                       "Value_Type", "Value_Description", "Value_ID", "X", "Y"], "", "MULTI_PART")
            arcpy.management.Delete("xtempmon")

            #                if int(arcpy.management.GetCount("tempmonlyr")[0]) > 0:  #if no sites detected, don't proceed further

            # Creating Phytopthora risk sites sites layer - add background poly, dissolve on required field
            print("   Processing Phytopthora sites data...")

            # One to one join to improve processing speed
            arcpy.analysis.SpatialJoin("DAP_buff1", "pchighrisk", "xtemppcrisk", join_operation="JOIN_ONE_TO_ONE",
                                       join_type="KEEP_COMMON", match_option="INTERSECT")
            arcpy.management.Dissolve("xtemppcrisk", "temppcrisk",
                                      ["DAP_REF_NO", "DAP_NAME", "DISTRICT", "DESCRIPTION", "RISK_LVL", "Value",
                                       "Value_Type"], "", "MULTI_PART")
            arcpy.management.Delete("xtemppcrisk")

            # Creating Ag Chemical Control Areas layer
            print("   Processing ACCA data...")

            arcpy.analysis.SpatialJoin("DAP_buff1", "acca", "xtempchem", join_operation="JOIN_ONE_TO_MANY",
                                       join_type="KEEP_ALL", match_option="INTERSECT")
            arcpy.management.Dissolve("xtempchem", "tempchem",
                                      ["DAP_REF_NO", "DAP_NAME", "DISTRICT", "DESCRIPTION", "RISK_LVL", "Value",
                                       "Value_Type"], "", "MULTI_PART")
            arcpy.management.Delete("xtempchem")

            print("   Processing Historic sites data...")

            arcpy.analysis.SpatialJoin("DAP_buff50", "HistHeritage", "xtemphist", join_operation="JOIN_ONE_TO_MANY",
                                       join_type="KEEP_ALL", match_option="INTERSECT")
            arcpy.management.Dissolve("xtemphist", "temphist",
                                      ["DAP_REF_NO", "DAP_NAME", "DISTRICT", "DESCRIPTION", "RISK_LVL", "Value",
                                       "Value_Type", "Value_Description", "Value_ID", "X", "Y"], "", "MULTI_PART")
            arcpy.management.Delete("xtemphist")

            ##                # Creating Historic sites layer from REGIONAL source data
            ##
            ##                print "   Processing 25k Historic sites data..."
            ##
            ##                arcpy.analysis.SpatialJoin("DAP_buff1", "HIST25_BUFF", "xtempreghis", join_operation="JOIN_ONE_TO_MANY",
            ##                           join_type="KEEP_ALL", match_option="INTERSECT")
            ##                arcpy.management.Dissolve("xtempreghis", "tempreghis", ["DAP_REF_NO","DAP_NAME","DISTRICT","DESCRIPTION","RISK_LVL","Value", "Value_Type", "X", "Y"], "", "MULTI_PART")
            ##                arcpy.management.Delete("xtempreghis")

            # Creating mining tenament layer from CSDL source data
            print("   Processing Mining tenament data...")
            arcpy.analysis.SpatialJoin("DAP_buff1", in_MIN, "xtempmin", join_operation="JOIN_ONE_TO_MANY",
                                       join_type="KEEP_ALL", match_option="INTERSECT")
            arcpy.management.AlterField("xtempmin", "TAG", "Value", "Value")
            arcpy.management.AlterField("xtempmin", "TYPEDESC", "Value_Type", "Value_Type")
            arcpy.management.Dissolve("xtempmin", "tempmin",
                                      ["DAP_REF_NO", "DAP_NAME", "DISTRICT", "DESCRIPTION", "RISK_LVL", "Value",
                                       "Value_Type"], "", "MULTI_PART")
            arcpy.management.Delete("xtempmin")

            # Creating mining site layer from CSDL source data
            print("   Processing Mine site data...")

            arcpy.analysis.SpatialJoin("DAP_buff50", "buffmsit", "xtempmsit", join_operation="JOIN_ONE_TO_MANY",
                                       join_type="KEEP_ALL", match_option="INTERSECT")
            arcpy.management.Dissolve("xtempmsit", "tempmsit",
                                      ["DAP_REF_NO", "DAP_NAME", "DISTRICT", "DESCRIPTION", "RISK_LVL", "Value",
                                       "Value_Type", "Value_ID", "X", "Y"], "", "MULTI_PART")
            arcpy.management.Delete("xtempmsit")

            # if int(arcpy.management.GetCount("tempmsit")[0]) > 0:

            # Creating Fire Operations Plan layer from CSDL source data
            ##                if mode <> "JFMP":
            print("   Processing JFMP data...")

            arcpy.analysis.SpatialJoin("DAP_buff1", "FOP_Gipps", "xtempfop", join_operation="JOIN_ONE_TO_MANY",
                                       join_type="KEEP_ALL", match_option="INTERSECT")
            arcpy.management.Dissolve("xtempfop", "tempfop",
                                      ["DAP_REF_NO", "DAP_NAME", "DISTRICT", "DESCRIPTION", "RISK_LVL", "Value",
                                       "Value_Type", "Value_Description"], "", "MULTI_PART")
            arcpy.management.Delete("xtempfop")

            # Creating Timber Release Plan layer from CSDL source data
            print("   Processing TRP data...")

            arcpy.analysis.SpatialJoin("DAP_buff1", "TRP", "xtemptrp", join_operation="JOIN_ONE_TO_MANY",
                                       join_type="KEEP_ALL", match_option="INTERSECT")
            arcpy.management.Dissolve("xtemptrp", "temptrp",
                                      ["DAP_REF_NO", "DAP_NAME", "DISTRICT", "DESCRIPTION", "RISK_LVL", "Value",
                                       "Value_Type"], "", "MULTI_PART")  # "Value_Description"
            arcpy.management.Delete("xtemptrp")

            # Creating apiary site layer from CSDL source data
            print("   Processing Apiary site data...")

            arcpy.analysis.SpatialJoin("DAP_buff1", "bee", "xtempbee", join_operation="JOIN_ONE_TO_MANY",
                                       join_type="KEEP_ALL", match_option="INTERSECT")
            arcpy.management.Dissolve("xtempbee", "tempbee",
                                      ["DAP_REF_NO", "DAP_NAME", "DISTRICT", "DESCRIPTION", "RISK_LVL", "Value",
                                       "Value_Type", "Value_ID", "X", "Y"], "", "MULTI_PART")
            arcpy.management.Delete("xtempbee")

            print("   Processing Utilities data...")
            arcpy.analysis.SpatialJoin("DAP_buff50", "Utilities", "xtemputil", join_operation="JOIN_ONE_TO_MANY",
                                       join_type="KEEP_ALL", match_option="INTERSECT")
            arcpy.management.Dissolve("xtemputil", "temputil",
                                      ["DAP_REF_NO", "DAP_NAME", "DISTRICT", "DESCRIPTION", "RISK_LVL", "Value",
                                       "Value_Type", "Value_Description"], "", "MULTI_PART")
            arcpy.management.Delete("xtemputil")

            print("   Processing railway data...")
            arcpy.analysis.SpatialJoin("DAP_buff50", "RAIL", "xtemprail", join_operation="JOIN_ONE_TO_MANY",
                                       join_type="KEEP_ALL", match_option="INTERSECT")
            arcpy.management.Dissolve("xtemprail", "temprail",
                                      ["DAP_REF_NO", "DAP_NAME", "DISTRICT", "DESCRIPTION", "RISK_LVL", "Value",
                                       "Value_Type", "Value_Description"], "", "MULTI_PART")
            arcpy.management.Delete("xtemprail")

            # Combine all the relevant layers from above with DAP layer...
            print("   Intersecting forest layers with DAP...")
            #  dapinfeat = ["temprainfor", "temphuts", "temprecweb", "tempgtrees", "tempog", "tempfmz"] #"temprainfor", "temphuts", "temprecweb", "tempgtrees", "tempog", "tempfmz", "tempreghis"
            infeat = ["tempfop", "temprecweb", "temppest", "tempmon", "tempchem", "temppcrisk", "temphist", "tempbee",
                      "tempmsit", "tempmin", "temptrp",
                      "tempgtrees", "tempog", "tempfmz", "temprainfor", "temphuts", "temprf", "tempglic", "temprail",
                      "temputil"]  # tempfop or Value recweb are the first because it has the largest character limit for value, value_type and value_desc. Place DAP values at very end
            featlist = []
            for feat in infeat:
                if arcpy.Exists(
                        feat):  # for the features that do exist, they will be added to a blank list for merging. This is for the differences in features available between LRLI and DAP values
                    featlist.append(feat)
            arcpy.management.Merge(featlist, "mergeforest")
            arcpy.management.MakeFeatureLayer("mergeforest", "DAP_theme")
            arcpy.management.SelectLayerByAttribute("DAP_theme", "NEW_SELECTION", "Value NOT IN('', NULL, 'NA')")

            print("Cleaning up temp covers")
            for fc in featlist:
                arcpy.management.Delete(fc)

            print("Forest values spatial check done")

            arcpy.analysis.Statistics("DAP_theme", "DAP_tab", [["OBJECTID", "COUNT"]],
                                      ["DAP_REF_NO", "DAP_NAME", "DISTRICT", "DESCRIPTION", "RISK_LVL", "Value_Type",
                                       "Value", "Value_Description", "Value_ID", "X",
                                       "Y"])  # "DAP_REF_NO", "DISTRICT", "DAP_NAME", "DESCRIPTION",
            print("Removed duplicate values")

            with arcpy.da.UpdateCursor("DAP_tab",
                                       ["X", "Y", "Value_Description"]) as cursor:  # missing XY values to have zeroes
                for row in cursor:
                    if (row[0] in (None, '')):
                        row[0] = 0
                    if (row[1] in (None, '')):
                        row[1] = 0
                    if (row[2] is None):
                        row[2] = ''
                    cursor.updateRow(row)

            # clean up
            for fc in ("mergeforest", "DAP_theme"):
                arcpy.management.Delete(fc)
            del featlist
            del infeat

        # ------------------------------------------------------------------------

        # BIODIVERSITY   -> intersecting all relevant input layers for biodiversity theme

        elif theme == "biodiversity":

            print("")
            print("  --> Creating combined layer for BIODIVERSITY table")

            # Creating temporary input layers from combined VBA source data

            if RiskLevel != "LRLI":

                # Creating clipped layer for Leadbeater possum sites
                print("   Clipping Leadbeaters Possum sites")
                if mode != 'JFMP':
                    arcpy.analysis.SpatialJoin("DAP_buff50", in_LBP, "xtemplbp", join_operation="JOIN_ONE_TO_MANY",
                                               join_type="KEEP_ALL", match_option="INTERSECT")
                else:
                    arcpy.analysis.SpatialJoin("JFMP_BioContingencyBuff", in_LBP, "xtemplbp",
                                               join_operation="JOIN_ONE_TO_MANY",
                                               join_type="KEEP_ALL", match_option="INTERSECT")
                arcpy.management.AlterField("xtemplbp", "LBP_DESC", "COMM_NAME", "COMM_NAME")
                arcpy.management.AddField("xtemplbp", "TYPE", "TEXT", 15)
                arcpy.management.CalculateField("xtemplbp", "TYPE", '"LBP Colony"', "PYTHON3")
                arcpy.management.Dissolve("xtemplbp", "templbp",
                                          ["DAP_REF_NO", "DAP_NAME", "DISTRICT", "DESCRIPTION", "RISK_LVL", "COMM_NAME",
                                           "TYPE", "BUFFER_TYPE"], "", "MULTI_PART")
                arcpy.management.Delete("xtemplbp")
                arcpy.management.AddField("templbp", "X", "DOUBLE", "", "0", "10")
                arcpy.management.AddField("templbp", "Y", "DOUBLE", "", "0", "10")
                with arcpy.da.UpdateCursor("templbp", ['SHAPE@X', 'SHAPE@Y', "X", "Y"],
                                           spatial_reference=srfilez) as cursor:
                    for row in cursor:
                        row[2] = int(row[0])
                        row[3] = int(row[1])
                        cursor.updateRow(row)

            # Creating clipped and merged layer for VBA sites
            Faunaexp = "(STARTDATE > date '1980-01-01 00:00:00') and (EPBC_DESC in ('Endangered','Vulnerable','Critically Endangered') or OLD_VICADV in ('cr','en','vu','wx') or FFG in ('cr','en','vu','th','cd','en-x','L') or COMM_NAME = 'Koala') and (MAX_ACC_KM <= 0.5)"
            Floraexp = "(STARTDATE > date '1980-01-01 00:00:00') and (EPBC_DESC in ('Endangered','Vulnerable','Critically Endangered') or OLD_VICADV in ('e','v','P','x') or FFG in('cr','en','vu','th','cd','en-x','L')) and (MAX_ACC_KM <= 0.5)"
            owlexp = "(STARTDATE > date '1980-01-01 00:00:00') and (COMM_NAME in ('Masked Owl','Powerful Owl','Sooty Owl', 'Square-tailed Kite','Barking Owl')) and (EXTRA_INFO in ('Roost site','Breeding'))"
            wbseexp = "(STARTDATE > date '1980-01-01 00:00:00') and (COMM_NAME = 'White-bellied Sea-Eagle') and (EXTRA_INFO in ('Roost site','Breeding'))"
            batexp = "(STARTDATE > date '1980-01-01 00:00:00') and (COMM_NAME in ('Grey-headed Flying-fox','Eastern Horseshoe Bat','Common Bent-wing Bat','Eastern Bent-winged Bat')) and (EXTRA_INFO in ('Roost site','Breeding'))"
            ghawkexp = "(STARTDATE > date '1980-01-01 00:00:00') and (COMM_NAME = 'Grey Goshawk') and (EXTRA_INFO in ('Roost site','Breeding'))"

            # if not arcpy.Exists("tempvba"): #redent the below if needed
            print("   Creating clipped VBA points")

            print("   Querying Flora data...")
            # Flora
            if mode != 'JFMP':
                # Initial data - vulnerable and endangered, and little 'r' listed

                arcpy.analysis.Clip(in_VBAFL25, "DAP_buff50", "xtempfl25")
                arcpy.analysis.Clip(in_VBAFLTHR, "DAP_buff50", "xtempflthr")
                arcpy.analysis.Clip(in_VBAFLRES, "DAP_buff50", "xtempflres")
                arcpy.management.Merge(["xtempfl25", "xtempflthr", "xtempflres"], "tempflora")

                arcpy.management.MakeFeatureLayer("tempflora", "floralyr")
                # Selecting rare flora species from a table list
                arcpy.management.SelectLayerByAttribute("floralyr", "NEW_SELECTION", rarelist_query)
                arcpy.management.SelectLayerByAttribute("floralyr", "ADD_TO_SELECTION", Floraexp)
                # arcpy.management.CopyFeatures("floralyr", "xtempfl25")
                arcpy.analysis.SpatialJoin("floralyr", "DAP_buff50", "tempfl25", join_operation="JOIN_ONE_TO_MANY",
                                           join_type="KEEP_ALL", match_option="INTERSECT")

            else:
                # Initial data - vulnerable and endangered, and little 'r' listed
                arcpy.analysis.Clip(in_VBAFL25, "JFMP_BioContingencyBuff", "xtempfl25")
                arcpy.analysis.Clip(in_VBAFLTHR, "JFMP_BioContingencyBuff", "xtempflthr")
                arcpy.analysis.Clip(in_VBAFLRES, "JFMP_BioContingencyBuff", "xtempflres")
                arcpy.management.Merge(["xtempfl25", "xtempflthr", "xtempflres"], "tempflora")

                arcpy.management.MakeFeatureLayer("tempflora", "floralyr")
                # Selecting rare flora species from a table list
                arcpy.management.SelectLayerByAttribute("floralyr", "NEW_SELECTION",rarelist_query)
                arcpy.management.SelectLayerByAttribute("floralyr", "ADD_TO_SELECTION", Floraexp)
                # arcpy.management.CopyFeatures("floralyr", "xtempfl25")
                arcpy.analysis.SpatialJoin("floralyr", "JFMP_BioContingencyBuff", "tempfl25",
                                           join_operation="JOIN_ONE_TO_MANY",
                                           join_type="KEEP_ALL", match_option="INTERSECT")

            arcpy.management.AddField("tempfl25", "TYPE", "TEXT", 15)
            arcpy.management.CalculateField("tempfl25", "TYPE", "'Flora'", "PYTHON3")

            print("   Querying Fauna data...")
            # Fauna
            if mode != 'JFMP':
                #  Initial clip uses the widest buffer, with further clipping when selections are done.
                arcpy.analysis.Clip(in_VBAFA25, "DAP_buff500", "xtempfa25")
                arcpy.analysis.Clip(in_VBAFATHR, "DAP_buff500", "xtempfathr")
                arcpy.analysis.Clip(in_VBAFARES, "DAP_buff500", "xtempfares")
                # arcpy.management.AlterField("xtempfathr", "OLD_VICADV", "VICADV", "VICADV") #field name recently changed in VBA_THREATENED datasets
                arcpy.management.Merge(["xtempfa25", "xtempfathr", "xtempfares"], "tempfauna")

                arcpy.management.MakeFeatureLayer("tempfauna", "faunalyr")

                # separating Sea Eagle records and re-buffering for owls, bats and goshawks
                arcpy.management.SelectLayerByAttribute("faunalyr", "NEW_SELECTION", wbseexp)
                arcpy.management.CopyFeatures("faunalyr", "xtempwbse")
                arcpy.analysis.SpatialJoin("xtempwbse", "DAP_buff500", "tempwbse", join_operation="JOIN_ONE_TO_MANY",
                                           join_type="KEEP_ALL", match_option="INTERSECT")

                arcpy.management.SelectLayerByAttribute("faunalyr", "NEW_SELECTION", owlexp)
                arcpy.management.CopyFeatures("faunalyr", "xtempowls")
                arcpy.analysis.Clip("xtempowls", "DAP_buff250", "ytempowls")
                arcpy.analysis.SpatialJoin("ytempowls", "DAP_buff250", "tempowls", join_operation="JOIN_ONE_TO_MANY",
                                           join_type="KEEP_ALL", match_option="INTERSECT")

                arcpy.management.SelectLayerByAttribute("faunalyr", "NEW_SELECTION", batexp)
                arcpy.management.CopyFeatures("faunalyr", "xtempbats")
                arcpy.analysis.Clip("xtempbats", "DAP_buff100", "ytempbats")
                arcpy.analysis.SpatialJoin("ytempbats", "DAP_buff100", "tempbats", join_operation="JOIN_ONE_TO_MANY",
                                           join_type="KEEP_ALL", match_option="INTERSECT")

                arcpy.management.SelectLayerByAttribute("faunalyr", "NEW_SELECTION", ghawkexp)
                arcpy.management.CopyFeatures("faunalyr", "xtempghawk")
                arcpy.analysis.Clip("xtempghawk", "DAP_buff250", "ytempghawk")
                arcpy.analysis.SpatialJoin("ytempghawk", "DAP_buff250", "tempghawk", join_operation="JOIN_ONE_TO_MANY",
                                           join_type="KEEP_ALL", match_option="INTERSECT")

                # Final buffering of all other fauna sites
                arcpy.management.SelectLayerByAttribute("faunalyr", "NEW_SELECTION", wbseexp)
                arcpy.management.SelectLayerByAttribute("faunalyr", "ADD_TO_SELECTION", owlexp)
                arcpy.management.SelectLayerByAttribute("faunalyr", "ADD_TO_SELECTION", batexp)
                arcpy.management.SelectLayerByAttribute("faunalyr", "ADD_TO_SELECTION", ghawkexp)
                arcpy.management.SelectLayerByAttribute("faunalyr", "SWITCH_SELECTION")
                arcpy.management.SelectLayerByAttribute("faunalyr", "SUBSET_SELECTION", Faunaexp)
                arcpy.management.SelectLayerByAttribute("faunalyr", "ADD_TO_SELECTION",
                                                        "(STARTDATE > date '1980-01-01 00:00:00') AND MAX_ACC_KM <= 0.5 AND TAXON_ID in (%s)" % rarelist)
                arcpy.management.CopyFeatures("faunalyr", "xtempothfa")
                arcpy.analysis.Clip("xtempothfa", "DAP_buff50", "ytempothfa")
                arcpy.analysis.SpatialJoin("ytempothfa", "DAP_buff50", "tempothfa", join_operation="JOIN_ONE_TO_MANY",
                                           join_type="KEEP_ALL", match_option="INTERSECT")

                #  Put it all together into a single fauna featureclass
                arcpy.management.Merge(["tempothfa", "tempowls", "tempwbse", "tempbats", "tempghawk"], "tempfa25")

            else:
                #  JFMP ONLY!
                arcpy.analysis.Clip(in_VBAFA25, "JFMP_BioContingencyBuff", "xtempfa25")
                arcpy.analysis.Clip(in_VBAFATHR, "JFMP_BioContingencyBuff", "xtempfathr")
                arcpy.analysis.Clip(in_VBAFARES, "JFMP_BioContingencyBuff", "xtempfares")
                # arcpy.management.AlterField("xtempfathr", "OLD_VICADV", "VICADV", "VICADV") #field name recently changed in VBA_THREATENED datasets
                arcpy.management.Merge(["xtempfa25", "xtempfathr", "xtempfares"], "tempfauna")

                arcpy.management.MakeFeatureLayer("tempfauna", "faunalyr")

                arcpy.management.SelectLayerByAttribute("faunalyr", "NEW_SELECTION", Faunaexp)
                arcpy.management.SelectLayerByAttribute("faunalyr", "ADD_TO_SELECTION",
                                                        "(STARTDATE > date '1980-01-01 00:00:00') AND MAX_ACC_KM <= 0.5 AND TAXON_ID in (%s)" % rarelist)
                # arcpy.management.CopyFeatures("faunalyr", "xtempothfa")
                arcpy.analysis.SpatialJoin("faunalyr", "JFMP_BioContingencyBuff", "tempfa25",
                                           join_operation="JOIN_ONE_TO_MANY",
                                           join_type="KEEP_ALL", match_option="INTERSECT")

            arcpy.management.AddField("tempfa25", "TYPE", "TEXT", 15)
            arcpy.management.CalculateField("tempfa25", "TYPE", "'Fauna'", "PYTHON3")

            #  Merge Flora and Fauna
            print("    Merging flora, fauna and threatened points")
            inlist = ["tempfl25", "tempfa25"]
            arcpy.management.Merge(inlist, "tempvba")

            for fc in "xtempfa25", "xtempfathr", "xtempowls", "xtempbats", "xtempghawk", "xtempothfa", "xtempwbse", "ytempowls", "ytempbats", "ytempghawk", "ytempothfa", "tempothfa", "tempowls", "tempwbse", "tempbats", "tempghawk", "tempfl25", "tempfa25", "tempflora", "floralyr", "faunalyr", "xtempfl25", "xtempflthr", "xtempflres":
                if arcpy.Exists(fc):
                    arcpy.management.Delete(fc)

            arcpy.management.DeleteIdentical("tempvba", ["DAP_REF_NO", "RECORD_ID"])  # remove any duplicates here

            #  Appending the results to a featureclass for future GIS work...
            print("    Appending resulting VBA sites to a GIS layer...")

            # arcpy.management.Project("tempvba", "VBA_outputs_z55", srfilez) #not needed?

            # adding xy coordinates to VBA table data to carry through the output
            arcpy.management.AddField("tempvba", "X", "DOUBLE", "", "0", "10")
            arcpy.management.AddField("tempvba", "Y", "DOUBLE", "", "0", "10")

            with arcpy.da.UpdateCursor("tempvba", ['SHAPE@X', 'SHAPE@Y', "X", "Y"],
                                       spatial_reference=srfilez) as cursor:
                for row in cursor:
                    row[2] = int(row[0])
                    row[3] = int(row[1])
                    cursor.updateRow(row)

            # table join TAXA LIST to grab TAXON_MOD field
            arcpy.management.JoinField("tempvba", "TAXON_ID", in_VBATaxaList, "TAXON_ID", ["TAXON_MOD"])

            # add VBA data into shapefile output, and remove any duplicates
            arcpy.management.Append("tempvba", "VBA_outputs", "NO_TEST")  # "VBA_outputs_z55"

            #   Buffering merged VBA points with 1 metre to convert to polys for the intersect
            #   This process also doubles as a dissolve on species using the field list option.
            arcpy.analysis.Buffer("tempvba", "tempvba_buff", "1 METER", "", "", "NONE",
                                  ["DAP_REF_NO", "DAP_NAME", "DISTRICT", "DESCRIPTION", "RISK_LVL", "SCI_NAME",
                                   "COMM_NAME", "TYPE",
                                   "TAXON_ID", "EXTRA_INFO", "RECORD_ID", "STARTDATE", "MAX_ACC_KM", "COLLECTOR",
                                   "FFG_DESC", "EPBC_DESC", "TAXON_MOD", "BUFFER_TYPE", "X",
                                   "Y"])  # LIST or NONE? #"tempvba" or #"VBA_outputs_z55"

            arcpy.management.MakeFeatureLayer("tempvba_buff", "vbalyr")  #
            arcpy.management.SelectLayerByAttribute("vbalyr", "NEW_SELECTION", "SCI_NAME = '' OR SCI_NAME IS NULL")
            arcpy.management.CalculateField("vbalyr", "SCI_NAME", '"NA"', "PYTHON3")
            arcpy.management.SelectLayerByAttribute("vbalyr", "NEW_SELECTION", "COMM_NAME = '' OR COMM_NAME IS NULL")
            arcpy.management.CalculateField("vbalyr", "COMM_NAME", '"NA"', "PYTHON3")
            arcpy.management.SelectLayerByAttribute("vbalyr", "NEW_SELECTION", "TAXON_ID IS NULL")
            arcpy.management.CalculateField("vbalyr", "TAXON_ID", '0', "PYTHON3")

            fclist = "tempfl25", "tempfl100", "tempfa25", "tempfa100", "tempvba", "vbalyr", "VBA_outputs_z55"
            for fc in fclist:
                arcpy.management.Delete(fc)
            del fclist

            arcpy.management.Rename("tempvba_buff", "tempvba")  # "tempvbastate"

            arcpy.management.AddField("tempvba", "ID_EXTRAINFO", "TEXT", 255)
            arcpy.management.CalculateField("tempvba", "ID_EXTRAINFO", "!DAP_REF_NO! + str(!TAXON_ID!)", "PYTHON3")
            arcpy.management.AddField("tempvba", "EXTRAINFO", "TEXT", 255)

            IDs = [row[0] for row in arcpy.da.SearchCursor("tempvba", ['ID_EXTRAINFO'], "TAXON_ID IS NOT NULL")]
            UniqueID = set(IDs)  # create unique list

            print("Flattening EXTRA_INFO")

            for ID in UniqueID:
                ref_exp = """{0} = '{1}'""".format(arcpy.AddFieldDelimiters("tempvba", 'ID_EXTRAINFO'), ID)
                VBA = [row[0] for row in arcpy.da.SearchCursor("tempvba", "EXTRA_INFO", ref_exp) if
                       row not in (None, '', 'None')]

                if len(VBA) >= 2:
                    UVBA = set(VBA)  # get unique value set
                    ValueVBA = ';'.join(
                        [str(item) for item in UVBA if item not in (None, '', 'None')])  # convert into a single string

                    with arcpy.da.UpdateCursor("tempvba", "EXTRAINFO", ref_exp) as cursor:
                        for row in cursor:
                            row[0] = str(ValueVBA)
                            cursor.updateRow(row)
                    del UVBA, ValueVBA

                del VBA

            # Check Species Recovery Overlay (SRO) data, for NBFT and JFMP only
            if "JFMP" in mode or "NBFT" in mode:
                print("    Checking Species Recovery Overlays...")
                arcpy.management.MakeFeatureLayer(in_SRO, "SROlyr")

                arcpy.management.SelectLayerByAttribute("SROlyr", "NEW_SELECTION", "SRO_GBC = 'Yes'")
                arcpy.analysis.SpatialJoin("SROlyr", "DAP_buff100", "tempGBC", join_operation="JOIN_ONE_TO_ONE",
                                           join_type="KEEP_COMMON", match_option="INTERSECT")
                arcpy.management.AlterField("tempGBC", "SRO_GBC", "VEG_CODE",
                                            "VEG_CODE")  # Use "VEG_CODE" field for later merging for data preparation in the Risk Register
                arcpy.management.CalculateField("tempGBC", "VEG_CODE", "'SROGBC'", "PYTHON3")

                arcpy.management.SelectLayerByAttribute("SROlyr", "NEW_SELECTION", "SRO_DPy = 'Yes'")
                arcpy.analysis.SpatialJoin("SROlyr", "DAP_buff100", "tempDPy", join_operation="JOIN_ONE_TO_ONE",
                                           join_type="KEEP_COMMON", match_option="INTERSECT")
                arcpy.management.AlterField("tempDPy", "SRO_DPy", "VEG_CODE", "VEG_CODE")
                arcpy.management.CalculateField("tempDPy", "VEG_CODE", "'SRODPy'", "PYTHON3")

                arcpy.management.SelectLayerByAttribute("SROlyr", "NEW_SELECTION", "SRO_SGG = 'Yes'")
                arcpy.analysis.SpatialJoin("SROlyr", "DAP_buff100", "tempSGG", join_operation="JOIN_ONE_TO_ONE",
                                           join_type="KEEP_COMMON", match_option="INTERSECT")
                arcpy.management.AlterField("tempSGG", "SRO_SGG", "VEG_CODE", "VEG_CODE")
                arcpy.management.CalculateField("tempSGG", "VEG_CODE", "'SROSGG'", "PYTHON3")

                arcpy.management.SelectLayerByAttribute("SROlyr", "NEW_SELECTION", "SRO_MO = 'Yes'")
                arcpy.analysis.SpatialJoin("SROlyr", "DAP_buff100", "tempMO", join_operation="JOIN_ONE_TO_ONE",
                                           join_type="KEEP_COMMON", match_option="INTERSECT")
                arcpy.management.AlterField("tempMO", "SRO_MO", "VEG_CODE", "VEG_CODE")
                arcpy.management.CalculateField("tempMO", "VEG_CODE", "'SROMO'", "PYTHON3")

                arcpy.management.SelectLayerByAttribute("SROlyr", "NEW_SELECTION", "SRO_SO = 'Yes'")
                arcpy.analysis.SpatialJoin("SROlyr", "DAP_buff100", "tempSO", join_operation="JOIN_ONE_TO_ONE",
                                           join_type="KEEP_COMMON", match_option="INTERSECT")
                arcpy.management.AlterField("tempSO", "SRO_SO", "VEG_CODE", "VEG_CODE")
                arcpy.management.CalculateField("tempSO", "VEG_CODE", "'SROSO'", "PYTHON3")

                arcpy.management.SelectLayerByAttribute("SROlyr", "NEW_SELECTION", "SRO_STQ = 'Yes'")
                arcpy.analysis.SpatialJoin("SROlyr", "DAP_buff100", "tempSTQ", join_operation="JOIN_ONE_TO_ONE",
                                           join_type="KEEP_COMMON", match_option="INTERSECT")
                arcpy.management.AlterField("tempSTQ", "SRO_STQ", "VEG_CODE", "VEG_CODE")
                arcpy.management.CalculateField("tempSTQ", "VEG_CODE", "'SROSTQ'", "PYTHON3")

                arcpy.management.Merge(["tempSTQ", "tempGBC", "tempDPy", "tempSGG", "tempMO", "tempSO"], "xtempSRO")

                arcpy.management.AddField("xtempSRO", "TYPE", "TEXT", 15)
                arcpy.management.CalculateField("xtempSRO", "TYPE", "'SRO'", "PYTHON3")

                arcpy.management.AddField("xtempSRO", "SCI_NAME", "TEXT", 255)
                arcpy.management.AddField("xtempSRO", "COMM_NAME", "TEXT", 255)

                with arcpy.da.UpdateCursor("xtempSRO", ["VEG_CODE", "SCI_NAME",
                                                        "COMM_NAME"]) as cursor:  # add in scientific name and common name
                    for row in cursor:
                        if row[0] == 'SROGBC':
                            row[1] = 'Calyptorhynchus lathami lathami'
                            row[2] = 'Glossy Black-Cockatoo'
                        if row[0] == 'SRODPy':
                            row[1] = 'Morelia spilota spilota'
                            row[2] = 'Diamond Python'
                        if row[0] == 'SROSGG':
                            row[1] = 'Petauroides volans'
                            row[2] = 'Greater Glider'
                        if row[0] == 'SROMO':
                            row[1] = 'Tyto novaehollandiae novaehollandiae'
                            row[2] = 'Masked Owl'
                        if row[0] == 'SROSO':
                            row[1] = 'Tyto tenebricosa tenebricosa'
                            row[2] = 'Sooty Owl'
                        if row[0] == 'SROSTQ':
                            row[1] = 'Dasyurus maculatus maculatus'
                            row[2] = 'Spot-tailed Quoll'
                        cursor.updateRow(row)

                arcpy.management.Dissolve("xtempSRO", "tempSRO",
                                          ["DAP_REF_NO", "DAP_NAME", "DISTRICT", "DESCRIPTION", "RISK_LVL", "COMM_NAME",
                                           "SCI_NAME", "TYPE", "VEG_CODE", "SCI_NAME", "COMM_NAME", "BUFFER_TYPE"])
                arcpy.management.DeleteIdentical("tempSRO", ["DAP_REF_NO", "VEG_CODE"])

            # Creating clipped layer for biodiversity monitoring sites
            # may use other monitoring site script? - this needs to be tested in greater detail
            print("    Intersecting Bio monitoring sites")

            if mode != 'JFMP':
                arcpy.analysis.SpatialJoin("DAP_buff50", "monflorafauna", "tempmsd", join_operation="JOIN_ONE_TO_MANY",
                                           join_type="KEEP_ALL", match_option="INTERSECT")
            else:
                arcpy.analysis.SpatialJoin("JFMP_BioContingencyBuff", "monflorafauna", "tempmsd",
                                           join_operation="JOIN_ONE_TO_MANY",
                                           join_type="KEEP_ALL", match_option="INTERSECT")
            ##                arcpy.management.MakeFeatureLayer ("tempmsd", "msdlyr")
            ##                arcpy.management.SelectLayerByAttribute ("msdlyr", "NEW_SELECTION", "COMM_NAME NOT IN('', NULL)") #MSD_NAME
            arcpy.management.Dissolve("tempmsd", "tempbiomonitor",
                                      ["DAP_REF_NO", "DAP_NAME", "DISTRICT", "DESCRIPTION", "RISK_LVL", "COMM_NAME",
                                       "SCI_NAME", "RECORD_ID", "TYPE", "BUFFER_TYPE", "X", "Y"], "", "MULTI_PART")
            arcpy.management.Delete("tempmsd")

            print("   Processing Rainforest data...")
            #                # Removing central Highlands areas from CSDL rainforest layer, will be using later reference_field data for this area.
            if not arcpy.Exists("RF_ALL"):
                print("   Creating one-off rainforest data..")
                arcpy.analysis.Erase(in_RFPOLY, in_RFCLIP, "xtemprfpoly")
                arcpy.management.Merge(["xtemprfpoly", in_RFPOLYCH], "RF_All")  # in_RFMURR]
                arcpy.management.MakeFeatureLayer("RF_All", "rflyr")
                arcpy.management.SelectLayerByAttribute("rflyr", "NEW_SELECTION", "RF = 1")
                arcpy.management.CalculateField("rflyr", "EVC_RF", '"Cool Temperate RF"',
                                                "PYTHON3")  # Central Highlands to be calculated as "cool temperature rainforest"
                # arcpy.management.SelectLayerByAttribute ("rflyr", "NEW_SELECTION", "ET_ID > 0")
                # arcpy.management.CalculateField("rflyr", "EVC_RF", '"Murrungowar RF"', "PYTHON3")
                arcpy.management.AlterField("RF_All", "EVC_RF", "COMM_NAME", "COMM_NAME")
                arcpy.management.AddField("RF_All", "TYPE", "TEXT", 15)
                arcpy.management.CalculateField("RF_All", "TYPE", '"Rainforest"', "PYTHON3")
                arcpy.management.AddField("RF_All", "VEG_CODE", "TEXT", 50)

                with arcpy.da.UpdateCursor("RF_All", ["VEG_CODE",
                                                      "COMM_NAME"]) as cursor:  # add in veg code (to be used as MR Code later on)
                    for row in cursor:
                        if row[1] == 'Littoral RF':
                            row[0] = 'LittoralRF'
                        if row[1] == 'Gallery RF':
                            row[0] = 'GalleryRF'
                        if row[1] == 'Dry RF':
                            row[0] = 'DryRF'
                        if row[1] == 'Warm Temperate RF':
                            row[0] = 'WarmTempRF'
                        if row[1] == 'Cool Temperate RF':
                            row[0] = 'CoolTempRF'
                        cursor.updateRow(row)

                arcpy.management.SelectLayerByAttribute("rflyr", "NEW_SELECTION", "RF = 0")  # remove world polygons
                if int(arcpy.management.GetCount("rflyr")[0]) > 0:
                    arcpy.management.DeleteFeatures("rflyr")

                for fc in "xtemprfpoly", "rflyr":
                    arcpy.management.Delete(fc)

                if mode != 'JFMP':
                    arcpy.analysis.SpatialJoin("DAP_buff50", "RF_ALL", "xtemprf", join_operation="JOIN_ONE_TO_MANY",
                                               join_type="KEEP_ALL", match_option="INTERSECT")
                else:
                    arcpy.analysis.SpatialJoin("JFMP_BioContingencyBuff", "RF_ALL", "xtemprf",
                                               join_operation="JOIN_ONE_TO_MANY",
                                               join_type="KEEP_ALL", match_option="INTERSECT")

                arcpy.management.Dissolve("xtemprf", "temprf",
                                          ["DAP_REF_NO", "DAP_NAME", "DISTRICT", "DESCRIPTION", "RISK_LVL", "COMM_NAME",
                                           "VEG_CODE", "BUFFER_TYPE", "TYPE"], "", "MULTI_PART")
                arcpy.management.Delete("xtemprf")

            #  Aquatic catchment layers
            print("    Intersecting Aquatic Catchment layer")
            if mode != 'JFMP':
                arcpy.analysis.SpatialJoin("DAP_buff1", in_AQUACATCH, "xtempaqua", join_operation="JOIN_ONE_TO_MANY",
                                           join_type="KEEP_ALL", match_option="INTERSECT")
            else:
                arcpy.analysis.SpatialJoin("JFMP_BioContingencyBuff", in_AQUACATCH, "xtempaqua",
                                           join_operation="JOIN_ONE_TO_MANY",
                                           join_type="KEEP_ALL", match_option="INTERSECT")

            arcpy.management.AlterField("xtempaqua", "GENUS", "TYPE")
            arcpy.management.CalculateField("xtempaqua", "TYPE", '!TYPE! + " Catchment"', "PYTHON3")
            arcpy.management.Dissolve("xtempaqua", "tempaqua",
                                      ["DAP_REF_NO", "DAP_NAME", "DISTRICT", "DESCRIPTION", "RISK_LVL", "COMM_NAME",
                                       "SCI_NAME", "TAXON_ID", "BUFFER_TYPE", "TYPE"], "", "MULTI_PART")
            arcpy.management.Delete("xtempaqua")

            #  Creating EVC inputs
            print("    Clipping EVCs")
            if mode != 'JFMP':
                arcpy.analysis.SpatialJoin("DAP_buff50", in_EVC, "xtempevc", join_operation="JOIN_ONE_TO_MANY",
                                           join_type="KEEP_ALL", match_option="INTERSECT")
            else:
                arcpy.analysis.SpatialJoin("JFMP_BioContingencyBuff", in_EVC, "xtempevc",
                                           join_operation="JOIN_ONE_TO_MANY",
                                           join_type="KEEP_ALL", match_option="INTERSECT")
            arcpy.management.AlterField("xtempevc", "X_EVCNAME", "COMM_NAME")
            arcpy.management.AddField("xtempevc", "TYPE", "TEXT", 15)
            arcpy.management.CalculateField("xtempevc", "TYPE", "'EVC'", "PYTHON3")
            arcpy.management.Dissolve("xtempevc", "tempevc",
                                      ["DAP_REF_NO", "DAP_NAME", "DISTRICT", "DESCRIPTION", "RISK_LVL", "COMM_NAME",
                                       "VEG_CODE", "TYPE", "BUFFER_TYPE"], "", "MULTI_PART")
            arcpy.management.Delete("xtempevc")

            # Combining all relevant input layers for biodiversity
            print("    Combining Biodiversity values with DAP")

            infeat = ["tempbiomonitor", "temprf", "tempevc", "tempvba", "templbp", "tempSRO",
                      "tempaqua"]  # "temprfad", "tempmsd",  "templbp" #"tempbiomonitor" to be first as it has the longest SCI_NAME field
            featlist = []
            for feat in infeat:
                if arcpy.Exists(
                        feat):  # for the features that do exist, they will be added to a blank list for merging. This is for the differences in features available between LRLI and DAP values
                    featlist.append(feat)
            arcpy.management.Merge(featlist, "DAP_theme")
            # arcpy.analysis.Intersect(["DAP_buff500","mergebiodiv"], "DAP_theme", "ALL", "0.0000001")

            # dapfclist = "templbp", "temprfad"
            print("Deleting temp covers")
            for fc in featlist:
                arcpy.management.Delete(fc)

            # Remove the background record where all fields are null
            arcpy.management.MakeFeatureLayer("DAP_theme", "frqlyr")
            arcpy.management.SelectLayerByAttribute("frqlyr", "NEW_SELECTION",
                                                    " \"COMM_NAME\" IN('NA', NULL, '') AND \"VEG_CODE\" IN('NA', NULL, '') AND \"TAXON_ID\" IN(0, NULL)")  # X_EVCNAME = 'NA' and SCI_FA25 = 'NA' and SCI_FL25 = 'NA' and X_RFCODE = 'NA' and MSD_NAME IN('NA', NULL)
            if int(arcpy.management.GetCount("frqlyr")[0]) > 0:
                arcpy.management.DeleteFeatures("frqlyr")
                print("    Null values removed")

            arcpy.management.Delete("frqlyr")
            arcpy.management.Delete("mergebiodiv")
            del featlist
            del infeat



        # ---------------------------------------------------------------------
        # WATER   -> intersecting all relevant input layers for water theme

        elif theme == "water":

            if RiskLevel != "LRLI":  # only run this theme for DAP activities where risk is not LRLI
                print("")
                print("  --> Creating combined layer for WATER table")
                # Creating a series of temporary layers converting points to polys and combining them
                # These layers will be deleted at the end of the script

                # Creating PWSC/Hydro layer
                print("   Processing catchments data")
                try:
                    arcpy.analysis.SpatialJoin("DAP_buff1", in_HYDRO, "temphydro", join_operation="JOIN_ONE_TO_ONE",
                                               join_type="KEEP_ALL", match_option="INTERSECT")
                    arcpy.management.MakeFeatureLayer("temphydro", "hydrolyr")
                    arcpy.management.SelectLayerByAttribute("hydrolyr", "NEW_SELECTION",
                                                            "FEATURE_TYPE_CODE <> '' OR FEATURE_TYPE_CODE IS NOT NULL")
                    arcpy.management.CalculateField("hydrolyr", "FEATURE_TYPE_CODE", '"Works on waterways"',
                                                    "PYTHON3")
                    arcpy.management.SelectLayerByAttribute("hydrolyr", "SWITCH_SELECTION")
                    arcpy.management.CalculateField("hydrolyr", "FEATURE_TYPE_CODE", '"Not on waterways"', "PYTHON3")
                    arcpy.management.SelectLayerByAttribute("hydrolyr", "NEW_SELECTION", "NAME IN ('', NULL)")
                    arcpy.management.CalculateField("hydrolyr", "NAME", '"NA"', "PYTHON3")
                    arcpy.management.Dissolve("temphydro", "temphydrod",
                                              ["DAP_REF_NO", "DAP_NAME", "DISTRICT", "DESCRIPTION", "NAME",
                                               "FEATURE_TYPE_CODE"], "", "MULTI_PART")

                    fclist = "temphydro", "hydrolyr"
                    for fc in fclist:
                        arcpy.management.Delete(fc)
                    arcpy.management.Rename("temphydrod", "temphydro")
                except:
                    print("Error in processing.  Running except clause")
                    if arcpy.Exists("temphydro"):
                        arcpy.management.Delete("temphydro")
                    arcpy.management.Copy("DAP_buff250", "temphydro")
                    arcpy.management.AddField("temphydro", "NAME", "TEXT", "15")
                    arcpy.management.MakeFeatureLayer("temphydro", "templyr")
                    arcpy.management.CalculateField("templyr", "NAME", '"OVERLAY ERROR"', "PYTHON3")
                    arcpy.management.Delete("templyr")

                print("   Processing Catchment Managment Authority boundaries")
                try:
                    arcpy.analysis.SpatialJoin("DAP_buff1", in_CMA, "xtempcma", join_operation="JOIN_ONE_TO_MANY",
                                               join_type="KEEP_ALL", match_option="INTERSECT")
                    arcpy.management.Dissolve("xtempcma", "tempcma",
                                              ["DAP_REF_NO", "DAP_NAME", "DISTRICT", "DESCRIPTION", "CMANAME"])
                    arcpy.management.AlterField("tempcma", "CMANAME", "CMA")
                    arcpy.management.Delete("xtempcma")

                except:
                    print("Error in processing.  Running except clause")
                    if arcpy.Exists("xtempcma"):
                        arcpy.management.Delete("xtempcma")
                    arcpy.management.Copy("DAP_buff250", "tempcma")
                    arcpy.management.AddField("tempcma", "CMA", "TEXT", "15")
                    arcpy.management.MakeFeatureLayer("tempcma", "templyr")
                    arcpy.management.CalculateField("templyr", "CMA", '"OVERLAY ERROR"', "PYTHON3")
                    arcpy.management.Delete("templyr")

                # Combine all the relevant layers from above with DAP layer...
                print("  Intersecting water layers with DAP...")
                if arcpy.Exists("DAP_theme"):
                    arcpy.management.Delete("DAP_theme")
                infeat = ["temphydro", "tempcma"]  # "DAP_buff1"
                arcpy.analysis.Intersect(infeat, "DAP_theme", "ALL", "0.0000001", "INPUT")
                del infeat

                print("Removing duplications...")
                # Summarise down to unique entries
                arcpy.analysis.Statistics("DAP_theme", "DAP_tab", [["OBJECTID", "COUNT"]],
                                          ["DAP_REF_NO", "DAP_NAME", "DISTRICT", "DESCRIPTION", "NAME",
                                           "FEATURE_TYPE_CODE", "CMA"])


        # ---------------------------------------------------------------------

        # HERITAGE   intersecting all relevant input layers for Heritage theme

        elif theme == "heritage":
            print("")
            print("  --> Creating combined layer for HERITAGE table")
            # Creating a series of temporary layers converting points to polys and combining them
            # These layers will be deleted at the end of the script

            print("   Getting Land Manager Data...")

            arcpy.analysis.SpatialJoin("DAP_buff1", in_PLM, "xtempplm", join_operation="JOIN_ONE_TO_MANY",
                                       join_type="KEEP_ALL", match_option="INTERSECT")
            arcpy.management.AddField("xtempplm", "LAND_MANGR", "TEXT", "255")

            with arcpy.da.UpdateCursor("xtempplm", "MNG_SPEC",
                                       "MNG_SPEC IN(NULL,'',' ')") as cursor:  # Blank land manager will imply it's privately owned
                for row in cursor:
                    row[0] = "Private"
                    cursor.updateRow(row)

            print("        Flattening PLM")
            IDs = [row[0] for row in arcpy.da.SearchCursor("xtempplm", 'DAP_REF_NO')]
            UniqueID = set(IDs)  # create unique list of DAP_REF_NO values

            for ID in UniqueID:
                ref_exp = """{0} = '{1}'""".format(arcpy.AddFieldDelimiters("xtempplm", "DAP_REF_NO"), ID)
                PLM = [row[0] for row in arcpy.da.SearchCursor("xtempplm", "MNG_SPEC", ref_exp) if
                       row not in (None, '', ' ', 'None')]

                if len(PLM) >= 2:  # checks if it has at least 2 objects in it, the loop will be bypassed if the list is empty or only has 1 entry
                    UPLM = set(PLM)  # get unique value set
                    ValuePLM = ';'.join([str(item) for item in UPLM])  # convert into a single string

                    with arcpy.da.UpdateCursor("xtempplm", "LAND_MANGR", ref_exp) as cursor:
                        for row in cursor:
                            row[0] = str(ValuePLM)
                            cursor.updateRow(row)
                    del UPLM, ValuePLM

                if len(PLM) == 1:
                    with arcpy.da.UpdateCursor("xtempplm", ["LAND_MANGR", "MNG_SPEC"], ref_exp) as cursor:
                        for row in cursor:
                            row[0] = row[1]
                            cursor.updateRow(row)

                del PLM

            arcpy.management.Dissolve("xtempplm", "tempplm",
                                      ["DAP_REF_NO", "DAP_NAME", "DISTRICT", "DESCRIPTION", "RISK_LVL", "AREA_HA",
                                       "Easting", "Northing", "YEAR_WORKS", "LAND_MANGR", "BUFFER_TYPE"])
            arcpy.management.Delete("xtempplm")

            # Process RAP and Sensitivity information
            print("   Processing RAP data")
            arcpy.analysis.SpatialJoin("DAP_buff1", in_RAP, "temprapx", join_operation="JOIN_ONE_TO_MANY",
                                       join_type="KEEP_ALL", match_option="INTERSECT")

            with arcpy.da.UpdateCursor("temprapx",
                                       "NAME") as cursor:  # rename/simplify RAP names here, "NAME IS NOT NULL"
                for row in cursor:
                    if (row[0] is None or row[0] == ''):
                        row[0] = "No RAP"
                    if ('Gunaikurnai Land and Waters Aboriginal Corporation' in row[0]):
                        row[0] = 'GLaWAC'
                    if ('Taungurung Land and Waters Council Aboriginal Corporation' in row[0]):
                        row[0] = 'Taungurung'
                    if ('Wurundjeri Woi Wurrung Cultural Heritage Aboriginal Corporation' in row[0]):
                        row[0] = 'Wurundjeri'
                    if ('Bunurong Land Council Aboriginal Corporation' in row[0]):
                        row[0] = 'Bunurong'
                    cursor.updateRow(row)

            print("        Flattening RAP")
            arcpy.management.AddField("temprapx", "CH_RAP", "TEXT", 255)

            # RAP Flatten
            IDs = [row[0] for row in arcpy.da.SearchCursor("temprapx", 'DAP_REF_NO')]
            UniqueID = set(IDs)  # create unique list of DAP_REF_NO values

            for ID in UniqueID:
                ref_exp = """{0} = '{1}'""".format(arcpy.AddFieldDelimiters("temprapx", "DAP_REF_NO"), ID)
                RAP = [row[0] for row in arcpy.da.SearchCursor("temprapx", "NAME", ref_exp) if
                       row not in (None, '', 'None')]

                if len(RAP) >= 2:  # checks if it has at least 2 objects in it, the loop will be bypassed if the list is empty or only has 1 entry
                    URAP = set(RAP)  # get unique value set
                    ValueRAP = ';'.join([str(item) for item in URAP])  # convert into a single string

                    with arcpy.da.UpdateCursor("temprapx", "CH_RAP", ref_exp) as cursor:
                        for row in cursor:
                            row[0] = str(ValueRAP)
                            cursor.updateRow(row)
                    del URAP, ValueRAP

                if len(RAP) == 1:
                    with arcpy.da.UpdateCursor("temprapx", ["CH_RAP", "NAME"], ref_exp) as cursor:
                        for row in cursor:
                            row[0] = row[1]
                            cursor.updateRow(row)

                del RAP

            arcpy.management.Dissolve("temprapx", "temprap",
                                      ["DAP_REF_NO", "DAP_NAME", "DISTRICT", "DESCRIPTION", "RISK_LVL", "Easting",
                                       "Northing", "YEAR_WORKS", "CH_RAP", "BUFFER_TYPE"], "", "MULTI_PART")
            arcpy.management.Delete("temprapx")

            print("   Processing SENSITIVITY data")
            arcpy.analysis.SpatialJoin("DAP_buff1", in_SENS, "tempsensx", join_operation="JOIN_ONE_TO_MANY",
                                       join_type="KEEP_ALL", match_option="INTERSECT")
            arcpy.management.AlterField("tempsensx", "SENSITIVITY", "CH_SENS", "CH_SENS")
            arcpy.management.CalculateField("tempsensx", "CH_SENS", "!CH_SENS!.title()", "PYTHON3")
            arcpy.management.Dissolve("tempsensx", "tempsens",
                                      ["DAP_REF_NO", "DAP_NAME", "DISTRICT", "DESCRIPTION", "RISK_LVL", "Easting",
                                       "Northing", "YEAR_WORKS", "CH_SENS", "BUFFER_TYPE"], "", "MULTI_PART")
            arcpy.management.Delete("tempsensx")
            with arcpy.da.UpdateCursor("tempsens", "CH_SENS", "CH_SENS IS NULL") as cur:
                for row in cur:
                    row[0] = "No"
                    cur.updateRow(row)

            print("   Processing Joint Management data")
            arcpy.analysis.SpatialJoin("DAP_buff1", in_JOINTMGMT, "tempjoinmgtx", join_operation="JOIN_ONE_TO_MANY",
                                       join_type="KEEP_ALL", match_option="INTERSECT")
            arcpy.management.AlterField("tempjoinmgtx", "LABELSHORT", "JointManagedPark", "JointManagedPark")

            print("        Flattening Joint Managed Park")

            for ID in UniqueID:
                ref_exp = """{0} = '{1}'""".format(arcpy.AddFieldDelimiters("tempjoinmgtx", "DAP_REF_NO"), ID)
                JMP = [row[0] for row in arcpy.da.SearchCursor("tempjoinmgtx", "JointManagedPark", ref_exp) if
                       row not in (None, '', 'None')]

                if len(JMP) >= 2:  # checks if it has at least 2 objects in it, the loop will be bypassed if the list is empty or only has 1 entry
                    UJMP = set(JMP)  # get unique value set
                    ValueJMP = ';'.join([str(item) for item in UJMP])  # convert into a single string

                    with arcpy.da.UpdateCursor("tempjoinmgtx", "JointManagedPark", ref_exp) as cursor:
                        for row in cursor:
                            row[0] = str(ValueJMP)
                            cursor.updateRow(row)
                    del UJMP, ValueJMP

                del JMP

            arcpy.management.Dissolve("tempjoinmgtx", "tempjoinmgt",
                                      ["DAP_REF_NO", "DAP_NAME", "DISTRICT", "DESCRIPTION", "RISK_LVL", "Easting",
                                       "Northing", "YEAR_WORKS", "JointManagedPark", "BUFFER_TYPE"], "", "MULTI_PART")
            arcpy.management.Delete("tempjoinmgtx")

            print("   Processing Parks Vic District data")
            arcpy.analysis.SpatialJoin("DAP_buff1", in_PV, "temppvx", join_operation="JOIN_ONE_TO_MANY",
                                       join_type="KEEP_ALL", match_option="INTERSECT")
            arcpy.management.AddField("temppvx", "PV_DISTRICT", "TEXT", "255")

            # Flatten to bring multiple PV districts into one field
            print("        Flattening PV District")

            IDs = [row[0] for row in arcpy.da.SearchCursor("temppvx", 'DAP_REF_NO')]
            UniqueID = set(IDs)  # create unique list of DAP_REF_NO values

            for ID in UniqueID:
                ref_exp = """{0} = '{1}'""".format(arcpy.AddFieldDelimiters("temppvx", "DAP_REF_NO"), ID)
                PVList = [row[0] for row in arcpy.da.SearchCursor("temppvx", "NAME", ref_exp)]

                if len(PVList) >= 2:  # checks if it has at least 2 objects in it, the loop will be bypassed if the list is empty or only has 1 entry
                    UPVList = set(PVList)  # get unique value set
                    ValuePV = ';'.join([str(item) for item in UPVList])  # convert into a single string

                    with arcpy.da.UpdateCursor("temppvx", "PV_DISTRICT", ref_exp) as cursor:
                        for row in cursor:
                            row[0] = str(ValuePV)
                            cursor.updateRow(row)
                    del UPVList, ValuePV

                if len(PVList) == 1:
                    with arcpy.da.UpdateCursor("temppvx", ["PV_DISTRICT", "NAME"], ref_exp) as cursor:
                        for row in cursor:
                            row[0] = row[1]
                            cursor.updateRow(row)

                del PVList

            arcpy.management.Dissolve("temppvx", "temppv",
                                      ["DAP_REF_NO", "DAP_NAME", "DISTRICT", "DESCRIPTION", "RISK_LVL", "Easting",
                                       "Northing", "YEAR_WORKS", "PV_DISTRICT", "BUFFER_TYPE"], "", "MULTI_PART")
            arcpy.management.Delete("temppvx")

            # Creating ACHRIS polygon layer from source site data

            if not arcpy.Exists("chsall_gda"):

                print("   Fetching ACHRIS sites data...")

                # modifying field names so it's easier to merge with preliminary sites
                arcpy.analysis.Clip(in_ACHRISS, "tempvic", "ACH_sites")
                arcpy.management.AlterField("ACH_sites", "COMPONENT_NO", "ACHRIS_ID", "ACHRIS_ID")
                arcpy.management.AlterField("ACH_sites", "DATE_MODIFIED", "LAST_MODIFIED", "LAST_MODIFIED")
                arcpy.management.AlterField("ACH_sites", "COMPONENT_TYPE", "PLACE_TYPE", "PLACE_TYPE")
                arcpy.management.AddField("ACH_sites", "STATUS", "TEXT", "100")
                arcpy.management.CalculateField("ACH_sites", "STATUS", '"Approved"', "PYTHON3")

                print("   Fetching ACHRIS preliminary reports data...")

                arcpy.analysis.Clip(in_ACHRISP, "tempvic", "PrelimClipped")
                arcpy.management.AlterField("PrelimClipped", "LOCATION", "PLACE_NAME", "PLACE_NAME")
                arcpy.management.AlterField("PrelimClipped", "PROJECT_NO", "ACHRIS_ID", "ACHRIS_ID")
                arcpy.management.AlterField("PrelimClipped", "DESCRIPTION", "PLACE_TYPE", "PLACE_TYPE")
                arcpy.management.AlterField("PrelimClipped", "RECORD_ADDED_DT", "LAST_MODIFIED", "LAST_MODIFIED")
                with arcpy.da.UpdateCursor("PrelimClipped", ['PLACE_TYPE']) as cursor:  # address None values
                    for row in cursor:
                        if row[0] is None:
                            row[0] = "NA"
                        cursor.updateRow(row)
                arcpy.management.MakeFeatureLayer("PrelimClipped", "PrelimSites",
                                                  "(STATUS IN('Complete' , 'Completed', 'Completed - not able to assess') OR STATUS LIKE 'Cur%' OR STATUS LIKE ('CUR%')) AND CURRENT_RECORD_YN = 'Y'")
                # arcpy.management.SelectLayerByAttribute("PrelimSites", "NEW_SELECTION", "(STATUS IN('Complete' , 'Completed', 'Completed - not able to assess') OR STATUS LIKE 'Cur%' OR STATUS LIKE ('CUR%')) AND CURRENT_RECORD_YN = 'Y'")
                # arcpy.management.Merge(["ACH_Sites", "PrelimSites"], "chsall")

                print("   Combining ACHRIS sites and preliminary reports data...")

                CHFields = ["SHAPE@", "ACHRIS_ID", "PLACE_NAME", "PLACE_TYPE", "STATUS", "LAST_MODIFIED"]

                with arcpy.da.SearchCursor("PrelimSites", CHFields) as sCur:
                    with arcpy.da.InsertCursor("ACH_sites", CHFields) as iCur:
                        for rows in sCur:
                            iCur.insertRow(rows)

                arcpy.management.Rename("ACH_sites", "chsall")
                arcpy.management.Project("chsall", "chsall_gda", srfilev) #Should we change it to srfilez?

                arcpy.management.AddField("chsall_gda", "X", "DOUBLE", "", "0", "10")
                arcpy.management.AddField("chsall_gda", "Y", "DOUBLE", "", "0", "10")

                with arcpy.da.UpdateCursor("chsall_gda", ['SHAPE@X', 'SHAPE@Y', "X", "Y"],
                                           spatial_reference=srfilez) as cursor:  # this adds XY data (easting/northing points in MGA Z55H) - uncomment this when ACHP_firesens and preliminary report on the CDSL are being used
                    for row in cursor:
                        row[2] = int(row[0])
                        row[3] = int(row[1])
                        cursor.updateRow(row)

                for fc in "PrelimSites", "chsall", "ACH_sites", "PrelimClipped":  # "chsallx" "chsall_buff550" "chsall_gda"
                    arcpy.management.Delete(fc)

                if mode != 'JFMP':
                    arcpy.analysis.Buffer("chsall_gda", "chsall_buff550", "550 meter", "FULL", "ROUND", "LIST",
                                          ["ACHRIS_ID", "PLACE_NAME", "PLACE_TYPE", "LAST_MODIFIED", "STATUS",
                                           "FIRE_SENSITIVITY", "X",
                                           "Y"])  # remove easting/northing from here when using data from CSD

            if mode == "JFMP":
                print("  Intersecting ACHRIS data with JFMP...")

                arcpy.analysis.SpatialJoin("JFMP_CHContingencyBuff", "chsall_gda", "tempchsx",
                                           join_operation="JOIN_ONE_TO_MANY",
                                           join_type="KEEP_ALL",
                                           match_option="INTERSECT")  # search_radius="5 Meters" ###Check if we go with WITHIN A DISTANCE 550m or just an intersect only?
                arcpy.management.Dissolve("tempchsx", "tempchsites",
                                          ["DAP_REF_NO", "DAP_NAME", "DISTRICT", "DESCRIPTION", "RISK_LVL", "Easting",
                                           "Northing", "YEAR_WORKS", "ACHRIS_ID", "PLACE_NAME", "PLACE_TYPE", "STATUS",
                                           "LAST_MODIFIED", "X", "Y", "BUFFER_TYPE", "FIRE_SENSITIVITY"], "",
                                          "MULTI_PART")
                arcpy.management.Delete("tempchsx")


            else:
                print("  Intersecting ACHRIS data with DAP...")

                arcpy.analysis.SpatialJoin("DAP_buff1", "chsall_buff550", "tempchsx", join_operation="JOIN_ONE_TO_MANY",
                                           join_type="KEEP_ALL", match_option="INTERSECT")  # search_radius="5 Meters"
                arcpy.management.Dissolve("tempchsx", "tempchsites",
                                          ["DAP_REF_NO", "DAP_NAME", "DISTRICT", "DESCRIPTION", "RISK_LVL", "Easting",
                                           "Northing", "YEAR_WORKS", "ACHRIS_ID", "PLACE_NAME", "PLACE_TYPE", "STATUS",
                                           "LAST_MODIFIED", "X", "Y", "BUFFER_TYPE", "FIRE_SENSITIVITY"], "",
                                          "MULTI_PART")
                arcpy.management.Delete("tempchsx")

            with arcpy.da.UpdateCursor("tempchsites", ["ACHRIS_ID", "STATUS"]) as cur:
                for row in cur:
                    if (row[1] not in (None, '', 'Approved')):
                        row[1] = "Preliminary"
                    if (row[0] is None or row[0] == ''):
                        row[0] = "NA"
                    cur.updateRow(row)

            # Combine all the relevant layers from above layer...
            print("  Intersecting heritage layers...")

            infeat = ["temprap", "tempsens", "tempplm", "tempjoinmgt", "temppv"]  # "tempchsites"
            # arcpy.Union_analysis(infeat, "DAP_themex", "ALL", "0.0000001") #this line to change?#########################################

            # combine all works level info together, appending through by JoinField
            arcpy.management.MakeTableView('tempplm', 'templmv')
            arcpy.management.MakeTableView('temppv', 'temppvv')
            arcpy.management.MakeTableView('tempjoinmgt', 'tempjoinmgtv')
            arcpy.management.MakeTableView('temprap', 'temprapv')
            arcpy.management.MakeTableView('tempsens', 'tempsensv')

            arcpy.management.JoinField("tempsensv", "DAP_REF_NO", 'templmv', "DAP_REF_NO", "LAND_MANGR")
            arcpy.management.JoinField("tempsensv", "DAP_REF_NO", 'temppvv', "DAP_REF_NO", "PV_DISTRICT")
            arcpy.management.JoinField("tempsensv", "DAP_REF_NO", 'tempjoinmgtv', "DAP_REF_NO", "JointManagedPark")
            arcpy.management.JoinField("tempsensv", "DAP_REF_NO", 'temprapv', "DAP_REF_NO", "CH_RAP")

            arcpy.conversion.TableToTable("tempsensv", workpath + "\\" + outGDB, "DAP_theme")

            ##                arcpy.management.Dissolve("DAP_themex", "DAP_theme", ["DAP_REF_NO", "DAP_NAME","DISTRICT","DESCRIPTION","RISK_LVL","Easting","Northing","YEAR_WORKS","ACHRIS_ID","PLACE_NAME","PLACE_TYPE",
            ##                                                                      "CH_RAP","CH_SENS","STATUS","JointManagedPark","LAST_MODIFIED","PV_DISTRICT","LAND_MANGR","X","Y", "BUFFER_TYPE", "FIRE_SENSITIVITY"], "", "MULTI_PART")
            ##
            ##                arcpy.management.Delete("DAP_themex")

            print("Cultural Heritage values spatial check done")

            del infeat

        # ---------------------------------------------------------------------

        # FOREST UTILISATION   -> intersecting all relevant input layers for Forest Utilisation theme
        # removed, most forutil values have been migrated to forests so it's all processed together

        #     TRANSFER OF VALUES FOUND IN THE OVERLAY PROCESS INTO OUTPUT TABLES
        # ===================================================================================================================================

        # Set up insert cursors to update tables...

        ##           if theme == "summary": #Obsolete as data flattening directly imports into Summary Table
        ##
        ##               if RiskLevel <> "LRLI":
        ##                   print "    Transferring values for Summary table"
        ##
        ##                   TabFields = ["DAP_REF_NO","DAP_NAME","DESCRIPTION","RISK_LVL","SCHEDULE", "EASTING", "NORTHING", "AREA_HA", "LENGTH_KM",
        ##                                "MMTGEN", "MNG_SPEC", "X_DESC", "ZONE_CODE", "OVERLAY", "LGA", "PLM_OVERLAY", "ALLOTMENT", "SEC", "PARISH_CODE", "P_NUMBER", "TOWNSHIP_CODE"] ## "DAP_REF_NO", "DAP_NAME", "DESCRIPTION",  #, "RN_NAME", "WZ_NAME", "NC_NAME", "HR_NAME", "RA_NAME" # "ALLOTMENT", "SEC", "PARISH_CODE", "P_NUMBER", "TOWNSHIP_CODE"
        ##                   with arcpy.da.SearchCursor("DAP_tab", TabFields) as sCur:
        ##                       with arcpy.da.InsertCursor(sumtab, ["UNIQUE_ID", "DISTRICT", "SITE_NAME", "SCHEDULE", "DESCRIPTION", "STATUS", "EASTING", "NORTHING", "AREA_HA", "LENGTH_KM", "LAND_STATUS",
        ##                                                            "LAND_MANGR", "FOR_TYPE", "PL_ZONE", "OVERLAY", "LGA", "PLM_OVERLAY", "C_ALLOT", "C_SEC", "C_PARISH", "C_PNUM", "C_TSHIP"]) as iCur:
        ##                           for rows in sCur:
        ##                    # Use insert cursor to add and populate table records - match the row number from DAP_tab in same sequence as the fields for the Summary Table
        ##                                iCur.insertRow((rows[0], str(dist), rows[1], rows[4], rows[2], rows[3], rows[5], rows[6], rows[7], rows[8], rows[9], rows[10], rows[11], rows[12], rows[13],
        ##                                                rows[14], rows[15], rows[16], rows[17], rows[18], rows[19], rows[20]))  ##  "DAP_REF_NO" = rows[0],"DAP_NAME" = rows[1],"DESCRIPTION" = rows[2], etc...

        if theme == "forests":
            print("   Transferring values for Forests table")
            # create an insert cursor to add and populate table records
            TabFields = ["DAP_REF_NO", "DAP_NAME", "DESCRIPTION", "Value_Type", "Value", "Value_Description", "X", "Y",
                         "RISK_LVL", "DISTRICT", "Value_ID"]  #

            with arcpy.da.SearchCursor("DAP_tab", TabFields) as sCur:
                with arcpy.da.InsertCursor(fortab, ["UNIQUE_ID", "DISTRICT", "NAME", "DESCRIPTION", "DATE_CHECKED",
                                                    "Value_Type", "Value", "Value_Description", "X", "Y", "RISK_LVL",
                                                    "Value_ID"]) as iCur:
                    for rows in sCur:
                        # Use insert cursor to add and populate table records - match the row number from DAP_tab in same sequence as the fields for the Values Summary Table
                        iCur.insertRow((rows[0], rows[9], rows[1], rows[2], start_date, rows[3], rows[4], rows[5],
                                        rows[6], rows[7], rows[8], rows[10]))  #

        # Biodiversity table values

        if theme == "biodiversity":
            print("   Transferring values for Biodiversity table")
            TabFields = ["SCI_NAME", "COMM_NAME", "EXTRAINFO", "TAXON_ID", "VEG_CODE", "TYPE", "X", "Y", "DAP_REF_NO",
                         "DAP_NAME", "DESCRIPTION", "RISK_LVL", "DISTRICT", "BUFFER_TYPE", "RECORD_ID", "STARTDATE",
                         "MAX_ACC_KM", "COLLECTOR", "FFG_DESC", "EPBC_DESC",
                         "TAXON_MOD"]  # fields to grab from DAP_theme

            with arcpy.da.SearchCursor("DAP_theme", TabFields) as sCur:
                with arcpy.da.InsertCursor(biotab,
                                           ["UNIQUE_ID", "DISTRICT", "NAME", "DESCRIPTION", "SCI_NAME", "COMM_NAME",
                                            "EXTRA_INFO", "TAXON_ID", "VEG_CODE", "TYPE", "X", "Y", "RISK_LVL",
                                            "BUFFER_TYPE", "RECORD_ID", "STARTDATE", "MAX_ACC_KM", "COLLECTOR",
                                            "FFG_DESC", "EPBC_DESC",
                                            "TAXON_MOD"]) as iCur:  # specify field names to populate from the blank template
                    for rows in sCur:
                        # Use insert cursor to add and populate table records - match the row number from DAP_theme in same order as the fields for the Biodiversity Table
                        iCur.insertRow((rows[8], rows[12], rows[9], rows[10], rows[0], rows[1], rows[2], rows[3],
                                        rows[4], rows[5], rows[6], rows[7], rows[11], rows[13], rows[14], rows[15],
                                        rows[16], rows[17], rows[18], rows[19], rows[20]))  #

            arcpy.management.CalculateField(biotab, "DATE_CHECKED", "(datetime.datetime.now()).strftime('%m/%d/%Y')",
                                            "python")

        # Water table values

        if theme == "water":

            if RiskLevel != "LRLI":
                print("   Transferring values for Water table")

                with arcpy.da.SearchCursor("DAP_tab",
                                           ["DAP_REF_NO", "DAP_NAME", "DESCRIPTION", "FEATURE_TYPE_CODE", "NAME", "CMA",
                                            "DISTRICT"]) as sCur:
                    with arcpy.da.InsertCursor(wattab, ["UNIQUE_ID", "DISTRICT", "NAME", "DESCRIPTION", "LU_DETERM",
                                                        "CATCH_NAME", "CMA"]) as iCur:
                        for rows in sCur:
                            iCur.insertRow((rows[0], rows[6], rows[1], rows[2], rows[3], rows[4], rows[5]))

        # Heritage table values TO EDIT
        if theme == "heritage":
            print("   Transferring values for Heritage tables")
            TabFields1 = ["DAP_REF_NO", "DISTRICT", "DAP_NAME", "DESCRIPTION", "ACHRIS_ID", "LAST_MODIFIED", "STATUS",
                          "PLACE_NAME", "PLACE_TYPE", "X", "Y",
                          "BUFFER_TYPE"]  # fields to grab from DAP_theme #"FIRE_SENSITIVITY"

            with arcpy.da.SearchCursor("tempchsites", TabFields1) as sCur:
                with arcpy.da.InsertCursor(hertab1, ["UNIQUE_ID", "DISTRICT", "NAME", "DESCRIPTION", "ACHRIS_ID",
                                                     "ACHRIS_DATEMODIFIED", "ACHRIS_STATUS", "PLACE_NAME", "PLACE_TYPE",
                                                     "X", "Y",
                                                     "BUFFER_TYPE"]) as iCur:  # specify field names  #"FIRE_SENSITIVITY"to populate from the blank template
                    for rows in sCur:
                        # Use insert cursor to add and populate table records - match the row number from DAP_theme in same order as the fields for the Heritage Table
                        iCur.insertRow((rows[0], rows[1], rows[2], rows[3], rows[4], rows[5], rows[6], rows[7], rows[8],
                                        rows[9], rows[10], rows[11]))  # ,rows[12]

            arcpy.management.CalculateField(hertab1, "SCRIPT_DATE", "(datetime.datetime.now()).strftime('%Y/%m/%d')",
                                            "python") #(datetime.datetime.now()).strftime('%m/%d/%Y')

            TabFields2 = ["DAP_REF_NO", "DISTRICT", "DAP_NAME", "RISK_LVL", "DESCRIPTION", "Easting", "Northing",
                          "YEAR_WORKS", "CH_RAP", "CH_SENS", "LAND_MANGR", "JointManagedPark", "PV_DISTRICT",
                          "BUFFER_TYPE"]

            with arcpy.da.SearchCursor("DAP_theme", TabFields2) as sCur:
                with arcpy.da.InsertCursor(hertab2,
                                           ["UNIQUE_ID", "DISTRICT", "NAME", "RISK_LVL", "DESCRIPTION", "Easting",
                                            "Northing", "Works_prog_yr", "CH_RAP", "CH_SENS", "LAND_MANGR",
                                            "JointManagedPark", "PV_District",
                                            "BUFFER_TYPE"]) as iCur:  # specify field names to populate from the blank template
                    for rows in sCur:
                        # Use insert cursor to add and populate table records - match the row number from DAP_theme in same order as the fields for the Heritage Table
                        iCur.insertRow((rows[0], rows[1], rows[2], rows[3], rows[4], rows[5], rows[6], rows[7], rows[8],
                                        rows[9], rows[10], rows[11], rows[12], rows[13]))  #

            arcpy.management.CalculateField(hertab2, "SCRIPT_DATE", "(datetime.datetime.now()).strftime('%Y/%m/%d')",
                                            "python") #(datetime.datetime.now()).strftime('%m/%d/%Y')

    print("")
    print("   Removing temp covers to prepare for the next entry...")

    fclist = ["DAP_buff50", "DAP_buff100", "DAP_buff250", "DAP_buff500", "DAP_temp", "DAP_theme", "tempvba_buff",
              "DAP_OSbuff1000", "DAP_OSbuff500", "tempchsites"]  # JFMP_CHContingencyBuff

    temps = (fc for fc in fclist if arcpy.Exists(fc))
    for fc in temps:
        arcpy.management.Delete(fc)

    del fclist

# ============================================== APPLY MITIGATIONS BELOW ===================================================================================================================

################### CALCULATE FORESTRY VALUES MITIGATION HERE - uses an update cursor to update the mitigation field according to the specific values found in Value_Type
# included here so it's done only once after the tables are complete
if int(arcpy.management.GetCount(fortab)[0]) > 0:
    print("")
    print("Applying mitigations to forest values...")

    with arcpy.da.UpdateCursor(fortab, ["Value_Type", "Mitigation"]) as cursor:
        for row in cursor:  # row[0] is value_type, row[1] is mitigation
            if (row[0] == 'Monitoring Site'):
                row[1] = "Prior to works, inform the contact listed on research site on the zoning layer"
            if (row[0] == 'PEST_PLANT'):
                row[
                    1] = "Ensure hygiene practices are adopted to reduce the risk of spreading weeds. Refer to Management Guideline - control of Pest Plants in the relevant FMP"
            if (row[0] == 'Agricultural Chemical Control Area'):
                row[
                    1] = "Chemical restrictions apply, this includes the type, timing and application method of herbicides. An ACCA permit will be required if works involve the use of herbicide. Check with Agriculture Victoria for guidance and additional information."
            if (row[0] == 'Phytophthora Risk'):
                row[
                    1] = r"Refer to risk mitigation strategies outlined in the Gippsland FMP or other relevant FMP and the Phytophthora cinnamomi SOP. Advice can be found in the Information Library in QuickBase by filtering Topics Covered to Hygiene"
            if (row[0] == 'Joint Fuel Management Plan'):
                row[
                    1] = "This activity coincides with a JFMP, engage with the district planned burn officer to confirm works do not conflict with burn prep area"
            if (row[0] == 'Mining Site'):
                row[
                    1] = "Ensure OH&S procedures are adopted to reduce the risk of injury. If an active site, notify licensee"
            if (row[0] == 'Mining Licence'):
                row[
                    1] = "Mining Licence present, ensure OH&S procedures are adopted to reduce the risk of injury. If an active site, notify licensee"
            if (row[0] == 'Apiary Site'):
                row[
                    1] = "Apiary site is susceptible to impacts from chemical use, timber harvesting and mature tree removal, fire and burning, removal of associated signage, and impacts to 2WD accessibility. Ensure works do not have a negative impact on apiary site. Licensee contact details can be found on Land Folio, contact licensee if your activity may cause impacts."
            if (row[0] == 'Historic Heritage Site'):
                row[
                    1] = "Refer to Management Guideline - Historic Places in the Relevant FMP. Any registered historic site (e.g. has number HXXXX or HXXXX-XXXX) may require a permit/consent prior to works, contact Heritage Victoria for additional guidance (VHR place permit - heritage.permits@delwp.vic.gov.au Heritage Inventory Archaeological Consent - archaeology.admin@delwp.vic.gov.au)"
            if (row[0] == 'TRP Coupe'):
                row[1] = "Works are near or overlap a VF coupe area. Engage with the relevant project planning team"
            if (row[0] == 'FMZ'):  # DAP values here and below
                row[
                    1] = "Refer to the Action Statement for the values identified or SMZ plan developed for the site if there is one. For more information refer to the Planning Standards for timber harvesting operations in Victoria's State Forests 2014"
            if (row[0] == 'Alpine Hut'):
                row[
                    1] = r"Engage with District Forest Planning Officer/Alpine Huts association/Parks Victoria where applicable. Refer to Management Guideline - Historic Places in the Relevant FMP"
            if (row[0] == 'Giant Tree'):
                row[1] = r"Engage with Regional NEP team. Especially for works involving native veg removal"
            ##            if (row[0] == 'Rainforest Site of Significance'):
            ##                row[1] = r"Engage with Regional NEP team. Works to remain within existing footprint, minimise any earthworks and spraying in protected areas containing RFSOS"
            if (row[0] in ('ASSET', 'CARPARK', 'HISTORIC RELIC', 'REC SITES', 'SIGNAGE')):  # RECWEB
                row[
                    1] = r"Recweb asset present. Engage with District Forest Planning Officer/Regional Recreation Planning Team. Refer to Management Guideline - Recreation Facility Management in the Relevant FMP"
            if (row[0] == 'Grazing Licence'):
                row[
                    1] = r"Engage with District Forest Planning Officer/ Regional Planning Officer responsible for Grazing"
            if (row[0] == 'Rainforest'):
                row[
                    1] = r"Engage with NEP team. Works to remain within existing footprint, minimise any earthworks and spraying near rainforests"
            ##            if (row[0] == 'Modelled Old Growth'):
            ##                row[1] = r"Old growth may occur here, minimise any earthworks and spraying near undisturbed, mature tree stands."
            if (row[0] == 'Powerline'):
                row[
                    1] = r"Works are in proximity to a powerline. Contact asset manager prior to works commencing if works are within 6.4m of the overhead power lines. Spotter from Ausnet services will be required if asset could be affected by works."
            if (row[0] == 'Pipeline'):
                row[
                    1] = r"Contact asset manager prior to works commencing if works are likely to impact the asset. No steel tracked machines to cross over gas pipeline unless designated cross over points are identified. No soil disturbance in proximity of gas pipeline. If location of pipeline is unknown or unclear use the online dial before you dig free service."
            if (row[0] == 'Railway'):
                row[
                    1] = r"Works are in proximity to a railway line. Contact asset manager prior to works commencing if works are likely to impact the asset. A permit may be required from VicTrack if works are on their land."
            cursor.updateRow(row)

    print("                                    ...Done")

################### Heritage Table post processing for a couple of fields (CH_RAPs and SITES_EXIST), included here so it's only done once.

if int(arcpy.management.GetCount(hertab2)[0]) > 0:

    # Calculate SITES_EXIST Field to note which activities have ACH sites within 500m

    ListIDsWithSites = ""
    with arcpy.da.SearchCursor(hertab1, 'UNIQUE_ID', "ACHRIS_ID <> 'NA'") as cursor:
        for row in cursor:
            ListIDsWithSites = ListIDsWithSites + "'" + str(row[0]) + "',"

    ListIDsWithSites = ListIDsWithSites[0:-1]

    arcpy.management.MakeTableView(hertab2, "hertab2view")

    try:
        arcpy.management.SelectLayerByAttribute("hertab2view", "NEW_SELECTION",
                                                "UNIQUE_ID IN(" + ListIDsWithSites + ")")
        arcpy.management.CalculateField("hertab2view", "SITES_EXIST", "'Yes'", "PYTHON3")
        arcpy.management.SelectLayerByAttribute("hertab2view", "SWITCH_SELECTION")
        if int(arcpy.management.GetCount("hertab2view")[0]) > 0:
            arcpy.management.CalculateField("hertab2view", "SITES_EXIST", "'No'", "PYTHON3")
    except:
        arcpy.management.SelectLayerByAttribute("hertab2view", "CLEAR_SELECTION")
        arcpy.management.CalculateField("hertab2view", "SITES_EXIST", "'No'", "PYTHON3")

    ################### CALCULATE HERITAGE VALUES MITIGATION HERE - uses an update cursor to update the mitigation field according by presence of sites, and cultural sensitivity layer
    # included here so it's done only once after the tables are complete

    print("Adding CH mitigations...")
    with arcpy.da.UpdateCursor(hertab2, ["SITES_EXIST", "CH_SENS", "MITIGATION", "RISK_LVL"]) as cursor:
        for row in cursor:  # row[0] is Site ID, row[1] is sensitivity, and row[2] is mitigation, row [3] is risk level
            if (row[3] == 'LRLI' and row[0] == 'No' and row[
                1] == 'No'):  # check for low impact works with no sites and no sensitivity
                row[
                    2] = "There are no known sites within the work area. Apply the contingency plan if any CH values is unearthed during proposed works. Remain within the existing footprint. No further assessment required. Continue works."
            if (row[3] == 'LRLI' and row[0] == 'No' and row[
                1] == 'Yes'):  # check for low impact works with no sites and has sensitivity
                row[
                    2] = "There are no known sites within the proposed work area. The proposed work area is within a Cultural sensitive area. Caution should be taken with machinery activity. Apply the contingency plan if any CH values is unearthed during proposed works. Remain within the footprint. No further assessment required. Continue works."
            if (row[3] != 'LRLI' and row[0] == 'No' and row[
                1] == 'No'):  # higher risk works with no sites or sensitivity
                row[
                    2] = "Works DO NOT intersect with known Aboriginal Places (sites) or areas of cultural heritage sensitivity. Ensure that all works remain within the disturbed footprint and are like for like replacement. The contingency plan for the discovery and management of heritage places should be on site and enacted in the event cultural heritage is found during works."
            if (row[3] != 'LRLI' and row[0] == 'No' and row[
                1] == 'Yes'):  # higher risk work with no sites but has sensitivity
                row[
                    2] = "Some or all of the proposed works are within a disturbed area that intersects an area of cultural heritage sensitivity. Ensure that all works remain within the disturbed footprint and are like for like replacement. The contingency plan for the discovery and management of heritage places should be on site and enacted in the event cultural heritage is found during works"
            ## #leave blank            if (row[0] == 'Yes' and row[1] == 'No'): #for works with sites and no sensitivity
            ## #                row[2] = "There are registered or preliminary Aboriginal Places (sites) within the area of works. If the works involve soil disturbance or vegetation removal, a check of ACHRIS is required by a Heritage Specialist or appropriately trained person to determine if site extent is known. If the site extent is UNKNOWN a site inspection is required by a Heritage Specialist to determine site extent and discuss mitigations. If the site extent is KNOWN, the Heritage Specialist will provide mitigation measures. If harm cannot be avoided, the District will need to apply for a permit in consultation with Heritage Specialist. As always ensure maintenance activities stay within disturbed footprint and the contingency plan is on site during works. Refer to QuickBase."
            ## #           if (row[0] == 'Yes' and row[1] == 'Yes'): #for works with sites and within sensitivity
            ## #               row[2] = "There are registered or preliminary Aboriginal Places (sites) AND areas of cultural heritage sensitivity within the area of works. If the works involve soil disturbance or vegetation removal, a check of ACHRIS is required by a Heritage Specialist or appropriately trained person to determine if site extent is known. If the site extent is UNKNOWN a site inspection is required by a Heritage Specialist to determine site extent and discuss mitigations. If the site extent is KNOWN, the Heritage Specialist will provide mitigation measures. If harm cannot be avoided, the District will need to apply for a permit in consultation with Heritage Specialist. As always ensure maintenance activities stay within disturbed footprint and the contingency plan is on site during works. Refer to QuickBase."

            cursor.updateRow(row)

    print("                      ...Done")

    #### summary statistics simplify down here for heritage tables hertab1 and hertab2, plus clean up

    fList1 = ['UNIQUE_ID', 'NAME', 'DISTRICT', 'DESCRIPTION', 'ACHRIS_ID', 'ACHRIS_DATEMODIFIED', 'ACHRIS_STATUS',
              'PLACE_TYPE', 'PLACE_NAME', 'X', 'Y', 'BUFFER_TYPE', 'FIRE_SENSITIVITY', "SCRIPT_DATE",
              "QB_ID"]  # [str(f.name) for f in arcpy.ListFields(hertab1) if f != 'OBJECTID']
    fList2 = ['UNIQUE_ID', 'NAME', 'DISTRICT', 'CH_SENS', 'MITIGATION', 'RISK_LVL', "Easting", "Northing",
              "Works_prog_yr", 'LAND_MANGR', 'JointManagedPark', 'BUFFER_TYPE', "PV_District", 'CH_RAP', 'GLAWAC_RAP',
              'TLAWC_RAP', 'WWWCHAC_RAP', 'BLCAC_RAP', 'NO_RAP', 'SITES_EXIST', "SCRIPT_DATE",
              "QB_ID"]  # [str(f.name) for f in arcpy.ListFields(hertab2) if f != 'OBJECTID'] #'DESCRIPTION', 'CH_RAP',

    print("Removing identical duplications for heritage tables... ")
    arcpy.analysis.Statistics(hertab1, "DAP_Heritage_SiteInfo", [["OBJECTID", "COUNT"]], fList1)
    arcpy.analysis.Statistics(hertab2, "DAP_Heritage_LandInfo", [["OBJECTID", "COUNT"]], fList2)

    with arcpy.da.UpdateCursor("DAP_Heritage_LandInfo", 'BUFFER_TYPE',
                               "BUFFER_TYPE IN('JFMP Contingency Area (500m from planned burn)', 'JFMP Contingency Buffer(500m from contingency area)')") as cursor:  # landinfo to be only provided at the JFMP Area extent
        for row in cursor:
            cursor.deleteRow()

    with arcpy.da.UpdateCursor("DAP_Heritage_SiteInfo", 'ACHRIS_ID',
                               "ACHRIS_ID IN('NA',NULL, '', ' ')") as cursor:  # works with no sites are removed
        for row in cursor:
            cursor.deleteRow()

    arcpy.management.DeleteField("DAP_Heritage_SiteInfo", "FREQUENCY")
    arcpy.management.DeleteField("DAP_Heritage_SiteInfo", "COUNT_OBJECTID")

    # Add LandInfo to Work Details table

    arcpy.management.JoinField("works_features", "DAP_REF_NO", "DAP_Heritage_LandInfo", "UNIQUE_ID",
                               ["CH_SENS", "MITIGATION", "LAND_MANGR", "JointManagedPark", "PV_District", "CH_RAP",
                                "SITES_EXIST", "SCRIPT_DATE"])

    # add to template for QB consistency
    arcpy.management.Append("works_features", worksfc, "NO_TEST")
else:  # if no heritage has been checked, works detail still appends to template, ignoring any blank fields
    arcpy.management.Append("works_features", worksfc, "NO_TEST")

############ NATIVE TITLE CALCULATION HERE
if (int(arcpy.management.GetCount(hertab2)[0]) > 0 and int(arcpy.management.GetCount(sumtab)[
                                                               0]) > 0):  # if summary table and heritage tables are populated, do the following
    arcpy.management.JoinField(sumtab, "UNIQUE_ID", worksfc, "DAP_REF_NO", ["CH_RAP"])  # Add RAP

    # Calculate NT_STATUS
    with arcpy.da.UpdateCursor(sumtab, ["NT_STATUS", "ACT"]) as cursor:
        for row in cursor:
            if row[0] == 'NO NT':
                row[0] = 'NT EXTINGUISHED'
            elif ('NO NT' in str(row[0]) and ('PART NT' in str(row[0]) or 'NT EXISTS' in str(row[0]))):
                row[0] = 'PART NT'
            elif (row[0] == 'PART NT' or 'PART NT' in str(row[0])):
                row[0] = 'PART NT'
            elif (row[1] == 'GOVERNMENT ROAD (AS TO PART)' or row[1] == 'GOVERNMENT ROAD'):
                row[0] = 'NT EXTINGUISHED'
            elif ('GOVERNMENT ROAD' in row[1] and row[0] in ('NT EXISTS', 'PART NT')):
                row[0] = 'PART NT'
            else:
                row[0] = 'NT EXISTS'
            cursor.updateRow(row)

    # Determine if there's permanent disturbance for activity and if activity is high impact
    arcpy.management.AddField(sumtab, "DisturbFlag", "TEXT", '', '',
                              '10')  # make an interim field to note if the activity involves permanent disturbance

    arcpy.management.AddField(sumtab, "ActivityFlag", "TEXT", '', '',
                              '10')  # make an interim field to note if the activity triggers Future Act Assessment

    High_Impact_worksList = ['Bridge-New', 'Bridge-Replacement', 'Culvert-New', 'Culvert-Replacement',
                             'Firebreak construction', 'New BBQ/Fire pit',
                             'New Shelters', 'New Toilet', 'Replace Bridge', 'Replace Ford with culvert or bridge',
                             'Road-New', 'Road widening/realignment']
    Maybe_Impact_worksList = ['Ford-New', 'Grazing', 'New Bike Track-by Machine', 'New Picnic Table',
                              'New Walking Track-by Machine']

    with arcpy.da.UpdateCursor(sumtab, ["DisturbFlag", "SOIL_DISTURB", "NATIVE_VEG", "ActivityFlag", "WORKS_TYPE_1",
                                        "WORKS_TYPE_2", "WORKS_TYPE_3"]) as cursor:
        for row in cursor:
            if ('permanent' in str(row[1]) or 'permanent' in str(row[2])):
                row[0] = 'Yes'
            else:
                row[0] = 'No'

            if (row[4] in High_Impact_worksList or row[5] in High_Impact_worksList or row[6] in High_Impact_worksList):
                row[3] = 'Yes'
            elif (row[4] in Maybe_Impact_worksList or row[5] in Maybe_Impact_worksList or row[
                6] in Maybe_Impact_worksList):
                row[3] = 'Maybe'
            else:
                row[3] = 'No'
            cursor.updateRow(row)

    # Determine if Reserved Land #TEST

    reserved = ['RESERVED FOREST (AS TO PART) [ACT NO. 6254/1958]', 'RESERVED FOREST [ACT NO. 6254/1958]',
                'CROWN LAND (RESERVED)',
                'CROWN LAND (RESERVED - AS TO PART)']  # pre set list of known Acts that signal reserved land
    unreserved = []  # empty list to populate later, to contain every other act in PLM25

    with arcpy.da.SearchCursor(in_PLM, ['ACT']) as cursor:
        for row in cursor:
            if (row[0] not in reserved and row[0] not in unreserved):
                unreserved.append(str(row[0]))
    arcpy.management.AddField(sumtab, "ReservedFlag", "TEXT", '', '',
                              '50')  # make an interim field to note if the activity is on reserved land, unreserved, or a mix of both

    with arcpy.da.UpdateCursor(sumtab, ["ReservedFlag", "ACT"]) as cursor:
        for row in cursor:

            act = str(row[1]).split('; ')
            res = 0
            unres = 0

            for t in act:
                if t in reserved:
                    res = 1
                elif t in unreserved:
                    unres = 1

            if res > 0 and unres == 0:  # check if only reserved
                row[0] = "Reserved"
            if res == 0 and unres > 0:  # check if only unreserved
                row[0] = "Unreserved"
            if res == 0 and unres == 0:  # check if there's no Act (default to unreserved)
                row[0] = "Unreserved"
            if res > 0 and unres > 0:  # check if there's both reserved and unreserved present
                row[0] = "Both Reserved and Unreserved"
            cursor.updateRow(row)

    with arcpy.da.UpdateCursor(sumtab, ["NT_Notification", "NT_STATUS", "CH_RAP", "ReservedFlag", "DisturbFlag",
                                        "ActivityFlag"]) as cursor:
        for row in cursor:
            if row[1] == 'NT EXTINGUISHED':  # check if extinguished
                row[0] = "Native Title Extinguished - No Procedural Rights Observed"
            elif (row[4] == 'No' or row[
                5] == 'No'):  # check if no permanent change to footprint and if works are low impact
                row[
                    0] = "Low Impact/Exempt Activity - Native Title Assessment Should Not Be Required Based On Provided Information"
            elif (row[5] == 'Maybe' or (row[3] != "Reserved" and row[4] == 'Yes' and row[5] == 'Yes') or (
                    row[1] == 'PART NT' and row[3] == "Reserved" and row[4] == 'Yes' and row[
                5] == 'Yes')):  # the 'maybe' or 'amber' clauses where more assessment is required
                row[0] = 'Seek Further Advice - Consult NT Assessor'
            elif (row[1] == 'NT EXISTS' and row[2] in ['GLaWAC'] and row[3] == "Reserved" and row[4] == 'Yes' and row[
                5] == 'Yes'):  # FAA applies within GLaWAC RAP
                row[
                    0] = "Future Act Procedural Rights Apply - 'Right To Comment'. Subdivision J - Construction of a 'Public Work'. Complete Future Act Notice, Gunaikurnai Notices (Determination Area)"
            elif (row[1] == 'NT EXISTS' and row[2] not in ['GLaWAC'] and row[3] == "Reserved" and row[4] == 'Yes' and
                  row[5] == 'Yes'):  # FAA applies outside GLaWAC RAP
                row[
                    0] = "Future Act Procedural Rights Apply - 'Right To Comment'. Subdivision J - Construction of a 'Public Work'. Complete Future Act Notice, FNLRS Notices (Non-Determined)"
            cursor.updateRow(row)

    # Remove temporary fields no longer needed in output
    flist = ["ReservedFlag", "DisturbFlag", "ActivityFlag"]

    for f in flist:
        arcpy.management.DeleteField(sumtab, f)

#################### JOIN MITIGATION RISK REGISTER TO BIODIVERSITY TABLE
# -------------------------------------------------------------------------------------------

if int(arcpy.management.GetCount(biotab)[0]) > 0:
    print("Setting up biodiversity risk table...")

    # Duplicate biodiversity table, add new field MR_CODE to be later used for joining

    arcpy.management.Copy(biotab, "risktemp")
    arcpy.management.AddField("risktemp", "MR_CODE", "TEXT", 50)
    arcpy.management.MakeTableView("risktemp", "risktempv")
    with arcpy.da.UpdateCursor("risktemp", ['X', 'Y'], "X IS NULL or Y IS NULL") as cursor:
        for row in cursor:
            row[0] = 0
            row[1] = 0
            cursor.updateRow(row)

    # Dissolve tables in 3 different ways (taxon - incl breeding and roosting, vegcode, biomonitoring), rename their MR_CODE and merge them together

    #### Summarise species data by using the Frequency tool, it may be necessary to retain multiple records of the same species in the same activity
    # arcpy.Frequency_analysis("risktempv", "taxontab", ["OBJECTID", "UNIQUE_ID", "DISTRICT", "NAME", "DESCRIPTION", "RISK_LVL", "Mitigation", "EXTRA_INFO", "MR_CODE", "SCI_NAME", "COMM_NAME", "TYPE", "X", "Y"], "TAXON_ID") #"FAU25_SPP", "FAU25_COM", "FLO25_SPP", "FLO25_COM"

    ### Summary Statistics now used to remove duplicate identical species with the same XY coordinates
    arcpy.analysis.Statistics("risktemp", "taxontab", [["OBJECTID", "COUNT"]],
                              ["UNIQUE_ID", "DISTRICT", "NAME", "DESCRIPTION", "RISK_LVL", "Mitigation", "DATE_CHECKED",
                               "EXTRA_INFO", "MR_CODE", "SCI_NAME", "COMM_NAME", "RECORD_ID", "STARTDATE", "MAX_ACC_KM",
                               "COLLECTOR", "FFG_DESC", "EPBC_DESC", "TAXON_MOD", "TYPE", "X", "Y", "TAXON_ID", "QB_ID",
                               "BUFFER_TYPE"])
    arcpy.management.MakeTableView("taxontab", "taxontabv")
    arcpy.management.SelectLayerByAttribute("taxontabv", "NEW_SELECTION",
                                            "(EXTRA_INFO LIKE '%Breeding%' OR EXTRA_INFO LIKE '%Roost site%') AND TAXON_ID IN(10220,10226,10230,10246,10248,10250,10253,11280,11303,61341,61342,10117,10118,1887,10176,10238,10177,10112,10215,60618,11141,10217,11455,10603,10277,10302,10268)")  # single out taxons with known breeding and roosting mitigations?
    # arcpy.management.SelectLayerByAttribute("taxontabv", "NEW_SELECTION", "(EXTRA_INFO LIKE '%Breeding%' OR EXTRA_INFO LIKE '%Roost site%')")
    if int(arcpy.management.GetCount("taxontabv")[0]) > 0:
        arcpy.management.CalculateField("taxontabv", "MR_CODE", "str(int(!TAXON_ID!)) + 'BR'",
                                        "PYTHON3")  # need to use str(int()) on taxon ID to ensure no decimals are carried over
    arcpy.management.SelectLayerByAttribute("taxontabv", "NEW_SELECTION", "TYPE = 'LBP Colony'")
    if int(arcpy.management.GetCount("taxontabv")[0]) > 0:
        arcpy.management.CalculateField("taxontabv", "MR_CODE", "'11141MON'", "PYTHON3")
        arcpy.management.CalculateField("taxontabv", "SCI_NAME", "'Gymnobelideus leadbeateri'", "PYTHON3")
    arcpy.management.SelectLayerByAttribute("taxontabv", "NEW_SELECTION", "TAXON_ID > 0 AND MR_CODE IS NULL")
    if int(arcpy.management.GetCount("taxontabv")[0]) > 0:
        arcpy.management.CalculateField("taxontabv", "MR_CODE", "!TAXON_ID!", "PYTHON3")

    print("Species table ready")

    #### EVC table dissolve with summary statistics, SRO will be treated the same here

    arcpy.analysis.Statistics("risktemp", "vegtab", [["OBJECTID", "COUNT"]],
                              ["UNIQUE_ID", "DISTRICT", "NAME", "DESCRIPTION", "RISK_LVL", "Mitigation", "DATE_CHECKED",
                               "VEG_CODE", "MR_CODE", "COMM_NAME", "SCI_NAME", "RECORD_ID", "STARTDATE", "MAX_ACC_KM",
                               "COLLECTOR", "FFG_DESC", "EPBC_DESC", "TAXON_MOD", "TYPE", "X", "Y", "QB_ID",
                               "BUFFER_TYPE"])  # "EVC"
    arcpy.management.MakeTableView("vegtab", "vegtabv")
    arcpy.management.SelectLayerByAttribute("vegtabv", "NEW_SELECTION",
                                            "TYPE IN('EVC', 'Rainforest', 'SRO')")  # VEG_CODE NOT IN('NA', NULL, '')
    if int(arcpy.management.GetCount("vegtabv")[0]) > 0:
        arcpy.management.CalculateField("vegtabv", "MR_CODE", "!VEG_CODE!", "PYTHON3")

    print("EVC table ready")

    #### Biodivesity monitoring table dissolve with summary statistics

    arcpy.management.SelectLayerByAttribute("risktempv", "NEW_SELECTION",
                                            "TYPE = 'Bio Monitoring'")  # COMM_NAME <> 'NA' AND VEG_CODE = 'NA' AND TAXON_ID = 0
    arcpy.analysis.Statistics("risktempv", "montab", [["OBJECTID", "COUNT"]],
                              ["UNIQUE_ID", "DISTRICT", "NAME", "DESCRIPTION", "RISK_LVL", "Mitigation", "DATE_CHECKED",
                               "COMM_NAME", "SCI_NAME", "MR_CODE", "RECORD_ID", "STARTDATE", "MAX_ACC_KM", "COLLECTOR",
                               "FFG_DESC", "EPBC_DESC", "TAXON_MOD", "X", "Y", "TYPE", "QB_ID", "BUFFER_TYPE"])
    arcpy.management.MakeTableView("montab", "montabv")
    arcpy.management.SelectLayerByAttribute("montabv", "NEW_SELECTION", "COMM_NAME NOT IN('NA', NULL, '')")
    if int(arcpy.management.GetCount("montabv")[0]) > 0:
        arcpy.management.CalculateField("montabv", "MR_CODE", "!COMM_NAME!", "PYTHON3")

    print("Biodiversity monitoring table ready")

    # something for Rainforest here?
    # arcpy.management.SelectLayerByAttribute("risktempv", "NEW_SELECTION", "TYPE IN ('Rainforest')")

    ### merge all the biodiversity tables together, removing any empty records

    arcpy.management.Merge(["taxontab", "vegtab", "montab"], "mergedtab")
    arcpy.management.MakeTableView("mergedtab", "mergedtabv")
    arcpy.management.SelectLayerByAttribute("mergedtabv", "NEW_SELECTION", "MR_CODE IS NULL")
    if int(arcpy.management.GetCount("mergedtabv")[0]) > 0:
        arcpy.management.DeleteRows("mergedtabv")

    print("Table set up complete")

    # --------------------------------------------------------------------------------------------
    ######## Begin applying the appropriate Risk Register
    # Join risk register table to the copy of Biodiv Summary table

    print("")
    print("Combining risk register to biodiveristy values...")

    if mode == "JFMP":
        if dist == "Latrobe":
            arcpy.management.Sort(JFMPriskreg, "riskreg", [
                ["Risk_Landscape", "DESCENDING"]])  # sort so EC values are at the top and are used first
        else:
            arcpy.management.Sort(JFMPriskreg, "riskreg", [
                ["Risk_Landscape", "ASCENDING"]])  # sort so AGG values are at the top and are used first

        arcpy.management.MakeTableView("mergedtab", "JFMPtab", "BUFFER_TYPE = 'JFMP Area'")
        arcpy.management.CopyRows("JFMPtab", "JFMPrisk")
        arcpy.management.MakeTableView("JFMPrisk", "JFMPtabv")
        arcpy.management.JoinField("JFMPtabv", "MR_CODE", "riskreg", "MR_Code",
                                   ["Threat", "Risk_Event", "Mitigation_Measure", "Last_Modified"])
        arcpy.management.CalculateField("JFMPtabv", "Mitigation", "!Mitigation_Measure!", "PYTHON3")

        # arcpy.RemoveJoin_management("mergedtab")

        arcpy.management.MakeTableView("mergedtab", "DAPtab", "BUFFER_TYPE = 'JFMP Contingency Area'")
        arcpy.management.CopyRows("DAPtab", "DAPrisk")
        arcpy.management.MakeTableView("DAPrisk", "DAPtabv")
        arcpy.management.JoinField("DAPtabv", "MR_CODE", JFMPriskreg, "MR_Code",
                                   ["Threat", "Risk_Event", "Mitigation_Measure", "Last_Modified"])
        arcpy.management.CalculateField("DAPtabv", "Mitigation", "!Mitigation_Measure!", "PYTHON3")

    else:
        arcpy.management.JoinField("mergedtabv", "MR_CODE", DAPriskreg, "MR_Code",
                                   ["Soil_Disturbance", "Veg_Alteration", "Waterway_Disturbance", "Chemical_Use",
                                    "Mitigation_Measure",
                                    "Last_Modified"])  # "Species__Scientific", "Common_Name", "Value_Type"

        arcpy.management.CalculateField("mergedtabv", "Mitigation", "!Mitigation_Measure!", "PYTHON3")
        arcpy.management.MakeTableView("mergedtabv", "DAPrisk")

    print("")
    print("Risk Register added to biodiversity values table")

    # -------------------------------------------------------------------------------------------

    # Cleaning up BioRiskValues Table, removing some columns, addressing blank mitigations and saving into the output .gdb

    with arcpy.da.UpdateCursor("DAPrisk", ["DISTRICT", "Mitigation", "RISK_LVL"]) as cursor:  # mergedtab?
        print("Tidying up biodiversity risk table...")
        for row in cursor:
            # flag values that haven't been addressed in the Risk Register
            if (row[1] is None):
                row[
                    1] = "No mitigation has been assigned for this value due to being unmatched in the Risk Register. Refer to NEP staff before proceeding."
            # The statements below are to remove EVCs that don't apply to the District - based on FMA boundaries (Applies to NBFTDAP RISK REGISTER)
            if (row[0] in ["Macalister", "Latrobe"] and 'This is a referred EVC in the East Gippsland FMA.' in row[1]):
                row[1] = "No comment in addition to general work practices and principles"
            if (row[0] == "Tambo" and 'This is a referred EVC in the Central Gippsland FMA.' in row[1]):
                row[1] = "No comment in addition to general work practices and principles"
            if (row[0] == "Snowy" and ('This is a referred EVC in the Central Gippsland FMA.' in row[
                1] or 'This is a referred EVC in the Tambo and Central Gippsland FMAs.' in row[
                                           1] or 'This is a referred EVC in the Tambo  FMA.' in row[1])):
                row[1] = "No comment in addition to general work practices and principles"
            # Adding more detail in LRLI EVC mitigation here (Applies to NBFTDAP RISK REGISTER)
            if (row[2] == "LRLI" and 'This is a referred EVC' in row[1]):
                row[
                    1] = "Work is within or adjacent to a referral EVC. Minimise soil disturbance; Minimise vegetation disturbance and alteration; Avoid new works outside the existing footprint; Minimise alteration to natural drainage patterns near or within the EVC; Minimise canopy disturbance and creation of gaps in canopy vegetation; Minimise chemical drift and off target spraying within the EVC; Minimise machinery movement within this EVC"
            cursor.updateRow(row)

    if arcpy.Exists("JFMPrisk"):
        with arcpy.da.UpdateCursor("JFMPrisk", ["DISTRICT", "Mitigation", "RISK_LVL"]) as cursor:  # mergedtab?
            print("Tidying up biodiversity risk table within JFMP area...")
            for row in cursor:
                # flag values that haven't been addressed in the Risk Register
                if (row[1] is None):
                    row[
                        1] = "No mitigation has been assigned for this value due to being unmatched in the Risk Register. Refer to NEP staff before proceeding."
                # The statements below are to remove EVCs that don't apply to the District - based on FMA boundaries (Applies to NBFTDAP RISK REGISTER)
                if (row[0] in ["Macalister", "Latrobe"] and 'This is a referred EVC in the East Gippsland FMA.' in row[
                    1]):
                    row[1] = "No comment in addition to general work practices and principles"
                if (row[0] == "Tambo" and 'This is a referred EVC in the Central Gippsland FMA.' in row[1]):
                    row[1] = "No comment in addition to general work practices and principles"
                if (row[0] == "Snowy" and ('This is a referred EVC in the Central Gippsland FMA.' in row[
                    1] or 'This is a referred EVC in the Tambo and Central Gippsland FMAs.' in row[
                                               1] or 'This is a referred EVC in the Tambo  FMA.' in row[1])):
                    row[1] = "No comment in addition to general work practices and principles"
                # Adding more detail in LRLI EVC mitigation here (Applies to NBFTDAP RISK REGISTER)
                if (row[2] == "LRLI" and 'This is a referred EVC' in row[1]):
                    row[
                        1] = "Work is within or adjacent to a referral EVC. Minimise soil disturbance; Minimise vegetation disturbance and alteration; Avoid new works outside the existing footprint; Minimise alteration to natural drainage patterns near or within the EVC; Minimise canopy disturbance and creation of gaps in canopy vegetation; Minimise chemical drift and off target spraying within the EVC; Minimise machinery movement within this EVC"
                cursor.updateRow(row)

    # Remove fields no longer needed in output
    flist = ["TAXON_ID", "FREQUENCY", "VEG_CODE", "COUNT_OBJECTID", "Mitigation_Measure"]  # "EXTRA_INFO"

    if mode != 'JFMP':
        for f in flist:
            arcpy.management.DeleteField("mergedtab", f)
        arcpy.management.Rename("mergedtab", "DAP_BiodiversityValues")
        biotable = "DAP_BiodiversityValues"

    else:
        for f in flist:
            arcpy.management.DeleteField("DAPrisk", f)
            arcpy.management.DeleteField("JFMPrisk", f)

        arcpy.management.Rename("DAPrisk", "JFMP_Contingency_BiodiversityValues")
        arcpy.management.Rename("JFMPrisk", "JFMP_BurnUnit_BiodiversityValues")
        biotable = "JFMP_BurnUnit_BiodiversityValues"

    # Add one more field to VBA_Outputs so that FMS number is visible (FMS Number aka BURNNO is populated by ASSET_ID in DAP template), also rename ASSET_ID in works_shapefile_template to FMS_ID
    if "JFMP" in mode or "NBFT" in mode:
        arcpy.management.AlterField(worksfc, "ASSET_ID", "FMS_ID", "FMS_ID")
        arcpy.management.JoinField("VBA_outputs", "DAP_REF_NO", worksfc, "DAP_REF_NO", ["FMS_ID"])

    # Another filter for EVC, if mitigation exists for the same EVC COMM_NAME, any duplicate EVC names against the same works starting with "No comment..." are removed
    # Dictionary to store UNIQUE_ID and corresponding EVC COMM_NAMEs that need to be deleted
    records_to_delete = {}

    # Using SearchCursor to iterate over the table
    with arcpy.da.SearchCursor(biotable, ['UNIQUE_ID', 'Mitigation', 'COMM_NAME'], "TYPE = 'EVC'") as cursor:
        for row in cursor:
            unique_id = row[0]
            mitigation = row[1]
            comm_name = row[2]

            # Check if there is a Mitigation other than 'No comment' for the same COMM_NAME and UNIQUE_ID
            if not mitigation.startswith("No comment"):
                if unique_id not in records_to_delete:
                    records_to_delete[unique_id] = set()
                records_to_delete[unique_id].add(comm_name)

    # Now delete records where the COMM_NAME has 'No comment' in the Mitigation field
    with arcpy.da.UpdateCursor(biotable, ['UNIQUE_ID', 'Mitigation', 'COMM_NAME'], "TYPE = 'EVC'") as cursor:
        for row in cursor:
            unique_id = row[0]
            mitigation = row[1]
            comm_name = row[2]

            # If the COMM_NAME for this UNIQUE_ID has a record with non-"No comment" mitigation,
            # and this record has "No comment", delete it
            if unique_id in records_to_delete and comm_name in records_to_delete[unique_id]:
                if mitigation.startswith("No comment"):
                    cursor.deleteRow()



# -------------------------------------FINAL CLEAN UP AND OUTPUT BELOW----------------------------------------------------------------

### Cleaning up remaining temp covers

print("")
print("Final clean up of temp covers")

featclass = arcpy.ListFeatureClasses()
tabs = arcpy.ListTables()

fc2keep = ["VBA_outputs", worksfc, "JFMP_CHContingencyBuff"]

temps = (fc for fc in featclass if fc not in fc2keep)
for fc in temps:
    arcpy.management.Delete(fc)

tabs2keep = [fortab, wattab, sumtab, "DAP_BiodiversityValues", "DAP_Heritage_SiteInfo", biofmQB, hertabQB,
             "JFMP_Contingency_BiodiversityValues", "JFMP_BurnUnit_BiodiversityValues"]  # hertab biotab

temps = (tab for tab in tabs if tab not in tabs2keep)
for tab in temps:
    arcpy.management.Delete(tab)

### Set up concatenated fields here for Unique Quickbase ID 'QB_ID' linkages, and Value_group for forest and biodiversity tables
if (arcpy.Exists("DAP_BiodiversityValues") and mode != 'NBFT'):
    with arcpy.da.UpdateCursor("DAP_BiodiversityValues",
                               ["QB_ID", "UNIQUE_ID", "COMM_NAME", "SCI_NAME", "MR_CODE", "X", "Y"]) as cursor:
        for row in cursor:
            if (row[3] is None or row[3] == ''):  # if SCI_NAME empty, don't use it (likely for EVC values only)
                row[0] = row[1] + '|' + row[2] + '|' + row[4] + '|' + str(int(row[5])) + '|' + str(int(row[6]))
            else:
                row[0] = row[1] + '|' + row[2] + ',' + row[3] + '|' + row[4] + '|' + str(int(row[5])) + '|' + str(
                    int(row[6]))
            cursor.updateRow(row)
    # arcpy.management.CalculateField("DAP_BiodiversityValues", "QB_ID", "!UNIQUE_ID! + '|' + !COMM_NAME! +  ',' + !SCI_NAME! + '|' + !MR_CODE! +'|' + str(int(!X!)) + '|' + str(int(!Y!))" ,"PYTHON3")
    arcpy.management.AddField("DAP_BiodiversityValues", "Value_Group", "TEXT", 50)
    arcpy.management.CalculateField("DAP_BiodiversityValues", "Value_Group", '"Biodiversity"', "PYTHON3")

    arcpy.management.AddField("DAP_BiodiversityValues", "QB_ID2", "TEXT",
                              255)  # add a secondary QBID using the newer format, for database comparison
    with arcpy.da.UpdateCursor("DAP_BiodiversityValues",
                               ["QB_ID2", "UNIQUE_ID", "RECORD_ID", "MR_CODE", "COMM_NAME", "TYPE"]) as cursor:
        for row in cursor:
            if (row[5] == "EVC"):
                row[0] = row[1] + '|' + row[3]
            elif (row[5] == "Bio Monitoring"):
                row[0] = row[1] + '|' + str(int(row[2])) + '|' + row[3]
            elif (row[5] in ["Flora", "Fauna"]):
                row[0] = row[1] + '|' + str(int(row[2]))
            elif (row[5] == "LBP Colony"):
                row[0] = row[1] + '|' + row[3] + '|' + row[4]
            else:
                row[0] = row[1] + '|' + row[3] + '|' + row[4] + '|' + row[5]
            cursor.updateRow(row)

if (arcpy.Exists("DAP_BiodiversityValues") and mode == 'NBFT'):
    with arcpy.da.UpdateCursor("DAP_BiodiversityValues",
                               ["QB_ID", "UNIQUE_ID", "RECORD_ID", "MR_CODE", "COMM_NAME", "TYPE", "SCI_NAME", "X",
                                "Y"]) as cursor:  # TEST with biomonitoring
        for row in cursor:
            if (row[5] == "EVC"):
                row[0] = row[1] + '|' + row[3]
            elif (row[5] == "Bio Monitoring"):
                row[0] = row[1] + '|' + str(int(row[2])) + '|' + row[3]
            elif (row[5] in ["Flora", "Fauna"]):
                # row[0] = row[1] + '|' + str(int(row[2])) #will revert back to old version
                row[0] = row[1] + '|' + row[4] + ',' + row[6] + '|' + row[3] + '|' + str(int(row[7])) + '|' + str(
                    int(row[8]))
            elif (row[5] == "LBP Colony"):
                row[0] = row[1] + '|' + row[3] + '|' + row[4]
            else:
                row[0] = row[1] + '|' + row[3] + '|' + row[4] + '|' + row[5]
            cursor.updateRow(row)
    arcpy.management.AddField("DAP_BiodiversityValues", "Value_Group", "TEXT", 50)
    arcpy.management.CalculateField("DAP_BiodiversityValues", "Value_Group", '"Biodiversity"', "PYTHON3")

if arcpy.Exists("JFMP_BurnUnit_BiodiversityValues"):
    with arcpy.da.UpdateCursor("JFMP_BurnUnit_BiodiversityValues",
                               ["QB_ID", "UNIQUE_ID", "RECORD_ID", "MR_CODE", "COMM_NAME", "TYPE", "SCI_NAME", "X",
                                "Y"]) as cursor:
        for row in cursor:
            if (row[5] == "EVC"):
                row[0] = row[1] + '|' + row[3]
            elif (row[5] == "Bio Monitoring"):
                row[0] = row[1] + '|' + str(int(row[2])) + '|' + row[3]
            elif (row[5] in ["Flora", "Fauna"]):
                # row[0] = row[1] + '|' + str(int(row[2])) #will revert back to old version
                row[0] = row[1] + '|' + row[4] + ',' + row[6] + '|' + row[3] + '|' + str(int(row[7])) + '|' + str(
                    int(row[8]))
            elif (row[5] == "LBP Colony"):
                row[0] = row[1] + '|' + row[3] + '|' + row[4]
            else:
                row[0] = row[1] + '|' + row[3] + '|' + row[4] + '|' + row[5]
            cursor.updateRow(row)

    with arcpy.da.UpdateCursor("JFMP_Contingency_BiodiversityValues",
                               ["QB_ID", "UNIQUE_ID", "RECORD_ID", "MR_CODE", "COMM_NAME", "TYPE", "SCI_NAME", "X",
                                "Y"]) as cursor:
        for row in cursor:
            if (row[5] == "EVC"):
                row[0] = row[1] + '|' + row[3]
            elif (row[5] == "Bio Monitoring"):
                row[0] = row[1] + '|' + str(int(row[2])) + '|' + row[3]
            elif (row[5] in ["Flora", "Fauna"]):
                # row[0] = row[1] + '|' + str(int(row[2]))
                row[0] = row[1] + '|' + row[4] + ',' + row[6] + '|' + row[3] + '|' + str(int(row[7])) + '|' + str(
                    int(row[8]))
            elif (row[5] == "LBP Colony"):
                row[0] = row[1] + '|' + row[3] + '|' + row[4]
            else:
                row[0] = row[1] + '|' + row[3] + '|' + row[4] + '|' + row[5]
            cursor.updateRow(row)

    arcpy.management.AddField("JFMP_BurnUnit_BiodiversityValues", "Value_Group", "TEXT", 50)
    arcpy.management.CalculateField("JFMP_BurnUnit_BiodiversityValues", "Value_Group", '"Biodiversity"', "PYTHON3")
    arcpy.management.AddField("JFMP_Contingency_BiodiversityValues", "Value_Group", "TEXT", 50)
    arcpy.management.CalculateField("JFMP_Contingency_BiodiversityValues", "Value_Group", '"Biodiversity"',
                                    "PYTHON3")

if int(arcpy.management.GetCount(fortab)[0]) > 0:
    with arcpy.da.UpdateCursor(fortab, ["QB_ID", "UNIQUE_ID", "Value", "Value_Description", "X",
                                        "Y"]) as cursor:  #####TO REVISE????? Change QB_ID calculation by forest value type??? TEST
        for row in cursor:
            if (row[3] is None or row[3] == ''):  # If values description empty, don't use it
                row[0] = row[1] + '|' + row[2] + '|' + str(int(row[4])) + '|' + str(int(row[5]))
            else:
                row[0] = row[1] + '|' + row[2] + '|' + row[3] + '|' + str(int(row[4])) + '|' + str(int(row[5]))
            cursor.updateRow(row)

    # Populate new QB_ID2 here, a temporary improved ID for Forest values
    with arcpy.da.UpdateCursor(fortab, ["QB_ID2", "UNIQUE_ID", "Value", "Value_Description", "X", "Y", "Value_Type",
                                        "Value_ID"]) as cursor:  #### TEST
        for row in cursor:
            if (row[6] == 'Monitoring Site'):  # ref + ID + value
                row[0] = row[1] + '|' + row[7] + '|' + row[2]
            elif (row[6] in ['Mining Site', 'Apiary Site', 'Historic Heritage Site'] and (
                    row[7] is not None)):  # ref + ID + type
                row[0] = row[1] + '|' + row[7] + '|' + row[6]
            elif (row[6] in ['ASSET', 'CARPARK', 'HISTORIC RELIC', 'REC SITES', 'SIGNAGE']):
                row[0] = row[1] + '|' + row[7] + '|RECWEB ' + row[6]
            elif (row[6] == 'FMZ'):
                row[0] = row[1] + '|' + row[7]

            elif (row[6] not in ['ASSET', 'CARPARK', 'HISTORIC RELIC', 'REC SITES', 'SIGNAGE', 'Mining Site',
                                 'Apiary Site', 'Monitoring Site'] and (row[3] is None or row[
                3] == '')):  # for all other values, and if values description empty, don't use it
                row[0] = row[1] + '|' + row[2] + '|' + str(int(row[4])) + '|' + str(int(row[5]))
            else:
                row[0] = row[1] + '|' + row[2] + '|' + row[3] + '|' + str(int(row[4])) + '|' + str(int(row[5]))
            cursor.updateRow(row)

    arcpy.management.AddField(fortab, "Value_Group", "TEXT", 50)
    arcpy.management.CalculateField(fortab, "Value_Group", '"Forest Management"', "PYTHON3")

# TEMPORARY MEASURE: Remove QB_ID duplicates and clean up records for biodiversity and forest tables before combining them together, return the latest record if duplicates are found
if arcpy.Exists("JFMP_BurnUnit_BiodiversityValues"):
    arcpy.management.Sort("JFMP_BurnUnit_BiodiversityValues", "JFMP_BurnUnit_BiodiversityValues_sorted",
                          [['STARTDATE', 'DESCENDING'], ['MR_CODE', 'ASCENDING']])
    arcpy.management.DeleteIdentical("JFMP_BurnUnit_BiodiversityValues_sorted", ["QB_ID"])
    arcpy.management.AddField("JFMP_BurnUnit_BiodiversityValues_sorted", "EVCtemp", "TEXT", field_length=5000)
    arcpy.management.CalculateField("JFMP_BurnUnit_BiodiversityValues_sorted", "EVCtemp",
                                    '!UNIQUE_ID! + !COMM_NAME! + !Mitigation!', "PYTHON3")
    arcpy.management.MakeTableView("JFMP_BurnUnit_BiodiversityValues_sorted",
                                   "JFMP_BurnUnit_BiodiversityValues_sorted2")
    arcpy.management.SelectLayerByAttribute("JFMP_BurnUnit_BiodiversityValues_sorted2", "NEW_SELECTION", "TYPE = 'EVC'")
    arcpy.management.DeleteIdentical("JFMP_BurnUnit_BiodiversityValues_sorted2", ["EVCtemp"])
    arcpy.management.DeleteField("JFMP_BurnUnit_BiodiversityValues_sorted2", "EVCtemp")

    arcpy.management.Delete("JFMP_BurnUnit_BiodiversityValues")
    arcpy.management.Rename("JFMP_BurnUnit_BiodiversityValues_sorted", "JFMP_BurnUnit_BiodiversityValues")

if arcpy.Exists("DAP_BiodiversityValues"):
    arcpy.management.Sort("DAP_BiodiversityValues", "DAP_BiodiversityValues_sorted",
                          [['STARTDATE', 'DESCENDING'], ['MR_CODE', 'ASCENDING']])
    arcpy.management.DeleteIdentical("DAP_BiodiversityValues_sorted", ["QB_ID"])
    arcpy.management.AddField("DAP_BiodiversityValues_sorted", "EVCtemp", "TEXT", field_length=5000)
    arcpy.management.CalculateField("DAP_BiodiversityValues_sorted", "EVCtemp",
                                    '!UNIQUE_ID! + !COMM_NAME! + !Mitigation!', "PYTHON3")
    arcpy.management.MakeTableView("DAP_BiodiversityValues_sorted", "DAP_BiodiversityValues_sorted2")
    arcpy.management.SelectLayerByAttribute("DAP_BiodiversityValues_sorted2", "NEW_SELECTION", "TYPE = 'EVC'")
    arcpy.management.DeleteIdentical("DAP_BiodiversityValues_sorted2", ["EVCtemp"])
    # arcpy.management.CalculateField("DAP_BiodiversityValues_sorted", "EVCtemp", '!UNIQUE_ID! + !COMM_NAME!',"PYTHON3")

    arcpy.management.DeleteField("DAP_BiodiversityValues_sorted2", "EVCtemp")
    # exit() ###TEST FOR GP-LTB-ERI-0261 and GP-LTB-SGD-0351 for no comment screening out duplicates, to test later

    ##    evcs = set()
    ##
    ##    with arcpy.da.UpdateCursor("DAP_BiodiversityValues_sorted2",["UNIQUE_ID", "COMM_NAME", "Mitigation"]) as cursor:
    ##        for evc, Mitigation in cursor:
    ##            if Mitigation.startswith("No comment"): #row[2]
    ##                if evc in evcs:
    ##                    cursor.deleteRow()
    ##                else:
    ##                    evcs.add(evc)
    ##                cursor.updateRow()
    # tabs2keep.append("DAP_BiodiversityValues_sorted")
    arcpy.management.Delete("DAP_BiodiversityValues")
    arcpy.management.Rename("DAP_BiodiversityValues_sorted", "DAP_BiodiversityValues")

if int(arcpy.management.GetCount(fortab)[0]) > 0:
    arcpy.management.Sort(fortab, "DAP_ForestValues", [['Value_ID', 'DESCENDING']])
    arcpy.management.DeleteIdentical("DAP_ForestValues", ["QB_ID"])
    arcpy.management.Delete(fortab)
    # fortab = "DAP_ForestValues"
    tabs2keep.append("DAP_ForestValues")

if arcpy.Exists("DAP_Heritage_SiteInfo"):  # subject to change soon for DAP
    if (mode == "JFMP" or mode == "NBFT"):
        arcpy.management.CalculateField("DAP_Heritage_SiteInfo", "QB_ID", "!UNIQUE_ID! + '|' + !ACHRIS_ID!",
                                        "PYTHON3")
    else:
        arcpy.management.CalculateField("DAP_Heritage_SiteInfo", "QB_ID",
                                        "!UNIQUE_ID! + '|' + !ACHRIS_ID! + '|'+ str(int(!X!)) + '|' + str(int(!Y!))",
                                        "PYTHON3")

### add tables to Quickbase readable format (combine forest management and biodiversity together, and Heritage Site and Land Info together)
if (mode == "JFMP"):  # because Risk Register gives us different fields to work with for JFMP vs DAP
    biofields = ['UNIQUE_ID', 'DISTRICT', 'NAME', 'DESCRIPTION', 'RISK_LVL', 'Mitigation', 'DATE_CHECKED', 'EXTRA_INFO',
                 'MR_CODE', 'SCI_NAME', 'COMM_NAME', 'TYPE', 'X', 'Y',
                 'QB_ID', 'Threat','Risk_Event','Last_Modified', 'Value_Group', "RECORD_ID", "STARTDATE", "MAX_ACC_KM", "COLLECTOR",
                 "FFG_DESC", "EPBC_DESC", "TAXON_MOD"]
    fmfields = ['UNIQUE_ID', 'DISTRICT', 'NAME', 'DESCRIPTION', 'RISK_LVL', 'Mitigation', 'DATE_CHECKED', 'Value_Type',
                'Value', 'Value_Description', 'X', 'Y', 'QB_ID', 'Value_Group', 'Value_ID']
    fmfields2 = ['UNIQUE_ID', 'DISTRICT', 'NAME', 'DESCRIPTION', 'RISK_LVL', 'Mitigation', 'DATE_CHECKED', 'TYPE',
                 'FM_Value', 'FM_Description', 'X', 'Y', 'QB_ID', 'Value_Group', 'RECORD_ID']  # input table has some fields named differently for forest management values

    if arcpy.Exists("JFMP_BurnUnit_BiodiversityValues"):
        arcpy.management.AddFields(biofmQB,[['Threat','TEXT','Threat',1000],['Risk_Event','TEXT','Risk_Event',3000]]) #last minute field additions for JFMP-exclusive risk register fields to carry over
        with arcpy.da.SearchCursor("JFMP_BurnUnit_BiodiversityValues", biofields) as sCur:
            with arcpy.da.InsertCursor(biofmQB,
                                       biofields) as iCur:  # specify field names to populate from the blank template
                for rows in sCur:
                    iCur.insertRow(rows)

    if arcpy.Exists("DAP_ForestValues"):  # if int(arcpy.management.GetCount("DAP_ForestValues")[0]) > 0:
        with arcpy.da.SearchCursor("DAP_ForestValues", fmfields) as sCur:
            with arcpy.da.InsertCursor(biofmQB,
                                       fmfields2) as iCur:  # specify field names to populate from the blank template
                for rows in sCur:
                    iCur.insertRow(rows)


# elif (mode == 'NBFT'):
else:
    biofields = ['UNIQUE_ID', 'DISTRICT', 'NAME', 'DESCRIPTION', 'RISK_LVL', 'Mitigation', 'DATE_CHECKED', 'EXTRA_INFO',
                 'MR_CODE', 'SCI_NAME', 'COMM_NAME', 'TYPE', 'X', 'Y',
                 'QB_ID', 'Soil_Disturbance', 'Veg_Alteration', 'Waterway_Disturbance', 'Chemical_Use', 'Last_Modified',
                 'Value_Group', "RECORD_ID", "STARTDATE", "MAX_ACC_KM", "COLLECTOR", "FFG_DESC", "EPBC_DESC",
                 "TAXON_MOD"]
    fmfields = ['UNIQUE_ID', 'DISTRICT', 'NAME', 'DESCRIPTION', 'RISK_LVL', 'Mitigation', 'DATE_CHECKED', 'Value_Type',
                'Value', 'Value_Description', 'X', 'Y', 'QB_ID', 'Value_Group', 'Value_ID']
    fmfields2 = ['UNIQUE_ID', 'DISTRICT', 'NAME', 'DESCRIPTION', 'RISK_LVL', 'Mitigation', 'DATE_CHECKED', 'TYPE',
                 'FM_Value', 'FM_Description', 'X', 'Y', 'QB_ID', 'Value_Group', 'RECORD_ID']  # input table has some fields named differently for forest management values

    if arcpy.Exists("DAP_BiodiversityValues"):
        with arcpy.da.SearchCursor("DAP_BiodiversityValues", biofields) as sCur:
            with arcpy.da.InsertCursor(biofmQB,
                                       biofields) as iCur:  # specify field names to populate from the blank template
                for rows in sCur:
                    iCur.insertRow(rows)

    if arcpy.Exists("DAP_ForestValues"):  # if int(arcpy.management.GetCount("DAP_ForestValues")[0]) > 0:
        with arcpy.da.SearchCursor("DAP_ForestValues", fmfields) as sCur:
            with arcpy.da.InsertCursor(biofmQB,
                                       fmfields2) as iCur:  # specify field names to populate from the blank template
                for rows in sCur:
                    iCur.insertRow(rows)

##else:
##    biofields = ['UNIQUE_ID','DISTRICT','NAME','DESCRIPTION','RISK_LVL','Mitigation','DATE_CHECKED','EXTRA_INFO','MR_CODE','SCI_NAME','COMM_NAME','TYPE','X','Y',
##                 'QB_ID','Soil_Disturbance','Veg_Alteration','Waterway_Disturbance','Chemical_Use','Last_Modified', 'Value_Group']
##    fmfields = ['UNIQUE_ID','DISTRICT','NAME','DESCRIPTION','RISK_LVL','Mitigation','DATE_CHECKED','Value_Type','Value','Value_Description','X','Y','QB_ID','Value_Group']
##    fmfields2 = ['UNIQUE_ID','DISTRICT','NAME','DESCRIPTION','RISK_LVL','Mitigation','DATE_CHECKED','TYPE','FM_Value','FM_Description','X','Y','QB_ID','Value_Group'] #input table has some fields named differently for forest management values
##
##
##    if arcpy.Exists("DAP_BiodiversityValues"):
##        with arcpy.da.SearchCursor("DAP_BiodiversityValues", biofields) as sCur:
##            with arcpy.da.InsertCursor(biofmQB, biofields) as iCur: #specify field names to populate from the blank template
##               for rows in sCur:
##                   iCur.insertRow(rows)
##
##    if arcpy.Exists("DAP_ForestValues"): #if int(arcpy.management.GetCount("DAP_ForestValues")[0]) > 0:
##        with arcpy.da.SearchCursor("DAP_ForestValues", fmfields) as sCur:
##            with arcpy.da.InsertCursor(biofmQB, fmfields2) as iCur: #specify field names to populate from the blank template
##               for rows in sCur:
##                   iCur.insertRow(rows)


##landfields = ['UNIQUE_ID','NAME','DISTRICT','CH_SENS','MITIGATION','RISK_LVL','Easting','Northing','Works_prog_yr','LAND_MANGR','JointManagedPark',
##              'PV_District','CH_RAP','GLAWAC_RAP','TLAWC_RAP','WWWCHAC_RAP','BLCAC_RAP','NO_RAP','SITES_EXIST','SCRIPT_DATE','QB_ID']
sitefields = ['UNIQUE_ID', 'NAME', 'DISTRICT', 'DESCRIPTION', 'ACHRIS_ID', 'ACHRIS_DATEMODIFIED', 'ACHRIS_STATUS',
              'PLACE_TYPE', 'PLACE_NAME', 'X', 'Y', 'BUFFER_TYPE', 'FIRE_SENSITIVITY', 'SCRIPT_DATE', 'QB_ID']

##CHQBfields1 = ['UNIQUE_ID','NAME','DISTRICT','CH_SENS','MITIGATION','RISK_LVL','Works_Easting','Works_Northing','Works_prog_yr','LAND_MANGR','JointManagedPark',
##              'PV_District','CH_RAP','GLAWAC_RAP','TLAWC_RAP','WWWCHAC_RAP','BLCAC_RAP','NO_RAP','SITES_EXIST','SCRIPT_DATE','QB_ID']
CHQBfields2 = ['UNIQUE_ID', 'NAME', 'DISTRICT', 'DESCRIPTION', 'ACHRIS_ID', 'ACHRIS_DATEMODIFIED', 'ACHRIS_STATUS',
               'PLACE_TYPE', 'PLACE_NAME', 'Site_X', 'Site_Y', 'BUFFER_TYPE', 'FIRE_SENSITIVITY', 'SCRIPT_DATE',
               'QB_ID']

if arcpy.Exists("DAP_Heritage_SiteInfo"):
    ##    with arcpy.da.SearchCursor("DAP_Heritage_LandInfo", landfields) as sCur:
    ##        with arcpy.da.InsertCursor(hertabQB, CHQBfields1) as iCur: #specify field names to populate from the blank template
    ##           for rows in sCur:
    ##               iCur.insertRow(rows)

    with arcpy.da.SearchCursor("DAP_Heritage_SiteInfo", sitefields) as sCur:
        with arcpy.da.InsertCursor(hertabQB,
                                   CHQBfields2) as iCur:  # specify field names to populate from the blank template
            for rows in sCur:
                iCur.insertRow(rows)

    arcpy.management.CalculateField(hertabQB, "IS90DAY", '"No"',
                                    "PYTHON3")  # Specify that this is not a 90 day check

### Write remaining tables into Spreadsheets and shapefiles, make sure to rename and move these spreadsheets outside of the workpath afterwards

tabkeepers = (tab for tab in tabs2keep if (arcpy.Exists(tab) and int(arcpy.management.GetCount(tab)[0]) > 0))
for excel in tabkeepers:
    arcpy.conversion.TableToExcel(excel,
                                  workpath + "\\" + str(start_date) + str(dist) + mode + "_" + str(excel)[4:] + ".xlsx",
                                  "ALIAS")
    print("   Saving " + workpath + "\\" + str(start_date) + str(dist) + mode + "_" + str(excel)[4:])

arcpy.conversion.TableToExcel(worksfc, workpath + "\\" + str(start_date) + str(dist) + mode + "_WorksDetail.xlsx",
                              "ALIAS")
print("   Saving " + workpath + "\\" + str(start_date) + str(dist) + mode + "_WorksDetail")

if arcpy.Exists("VBA_outputs"):
    arcpy.conversion.FeatureClassToFeatureClass("VBA_outputs", workpath,
                                                str(start_date) + str(dist) + mode + "_VBA_Outputs.shp")
    print("   Saving " + workpath + "\\" + str(start_date) + str(dist) + mode + "_VBA_Outputs")

### Save works as individual shapefiles to be zipped and uploaded onto ECM for later
if output_shapefile == True:
    print("Making individual shapefiles...")
    try:
        if mode == 'JFMP':
            arcpy.analysis.SplitByAttributes('JFMP_CHContingencyBuff', shapefilepath, ["DAP_REF_NO", "DAP_NAME"])
        else:
            arcpy.analysis.SplitByAttributes(worksfc, shapefilepath, ["DAP_REF_NO", "DAP_NAME"])
        print("Saved in " + shapefilepath)
    except:
        print("No shapefiles generated. Had an issue saving to: " + shapefilepath)

# convert .xls to .csv here
for xls in glob.iglob(os.path.join(workpath, '*.xlsx')):
    #xls = str(start_date) + str(dist) + mode + "_" + str(excel)[4:] + ".xlsx"
    read_file = pd.read_excel(xls)
    read_file.to_csv(xls[:-5] + '.csv', index=False, header=True, encoding='utf-8')
    os.remove(xls)
    #xls2 = str(start_date) + str(dist) + mode + "_WorksDetail.xlsx"
    #read_file = pd.read_excel(xls2)

#read_file.to_csv(xls2[:-5] + '.csv', header=True, encoding='utf-8')
# if mode != "JFMP" or mode != "NBFT":
#     for csv in glob.iglob(os.path.join(workpath, '*.csv')):
#
#         # Step 1: Authenticate using a service account
#         SERVICE_ACCOUNT_FILE = workpath + "/solid-future-442523-b7-a7da8fa28be8.json"  # Replace with your key file path
#         SCOPES = ['https://www.googleapis.com/auth/drive.file']
#         credentials = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
#
#         # Step 2: Generate an access token
#         access_token = credentials.token
#
#         if not access_token:  # Manually refresh if needed
#             credentials.refresh(Request())
#             access_token = credentials.token
#
#         # Step 3: Specify the file and folder information
#         # check CSV file and assign Folder ID
#         if '_WorksDetail' in csv:
#             folder_id = "1QsZ70vAEO3Z8TvMO_tkUB9shnjLzr5eB" #Raw Data describing works.
#         elif '_Summary_' in csv:
#             folder_id = '1oFv7Eh21b06DyvIXiG0w6rNhWZ3hml1j' #Native Title
#         # elif '' in csv:
#         #     folder_id = '1rOJW5M_Ry2XnqXwoTbKpxikRtfnV-Ygu' # NEP Advice
#         elif '_Biodiversity_Forests' in csv:
#             folder_id = '1-B6YQVqCPheDqGIIiIjJF3pbgvlCj-mN' #Biodiversity and forest values Advice
#         elif '_Heritage_SiteInfoQB' in csv:
#             folder_id = '1Wd3t7kQWKQiM7Dzx4wOe1JB77hb_YI96'  # Heritage Site Info
#         else:
#             print('folder id set to parent') # Parent Folder 'QuickBase Sync'
#             folder_id = '1g_6BPK3yQD0j7sWC7yG0xBYAkeVnTqtj'
#
#         # Step 4: Upload the file to Google Drive
#         url = 'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart'
#         headers = {"Authorization": f"Bearer {access_token}"}
#
#         # Prepare metadata and file content
#         metadata = {
#             "name": csv,  # The file's name on Drive
#             "parents": [folder_id],       # Upload to this folder
#         }
#
#         files = {
#             "metadata": ("metadata", json.dumps(metadata), "application/json"),
#             "file": (csv, open(csv, "rb")),  # File content
#         }
#
#         response = requests.post(url, headers=headers, files=files)
#
#         # Check the result
#         if response.status_code == 200:
#             print(f"File uploaded successfully: {response.json()['id']}")
#         else:
#             print(f"Failed to upload file: {response.content}")

### Direct Connection to QuickBase - Future Development Oppertunity
# dictionary_map_NT = {'UNIQUE_ID': 26, 'ACT': 6, 'AREA_HA': 7, 'C_ALLOT': 8, 'C_PARISH': 9, 'C_PNUM': 10, 'C_SEC': 11, 'C_TSHIP': 12, 'EASTING': 13, 'FOR_TYPE': 14, 'LAND_MANGR': 15, 'LAND_STATUS': 16, 'LENGTH_KM': 17, 'LGA': 18, 'NORTHING': 19, 'NT_Name': 20, 'NT_Notification': 21, 'NT_STATUS': 22, 'OVERLAY': 23, 'PLMOVERLAY': 24, 'PLZONE': 25, 'LIVE DB - Overlay Code': 27, 'LIVE DB - Parcel Number': 28, 'LIVE DB - Parish': 29, 'LIVE DB - ACT': 30, 'LIVE DB - Area (Ha)': 31, 'LIVE DB - Crown Allotment': 32, 'LIVE DB - Easting': 33, 'LIVE DB - Forest Type': 34, 'LIVE DB - Land Manager': 35, 'LIVE DB - Land Status': 36, 'LIVE DB - Length (km)': 37, 'LIVE DB - Local Government Area': 38, 'LIVE DB - Native Title Name': 39, 'LIVE DB - Native Title Status': 40, 'LIVE DB - Northing': 41, 'LIVE DB - NT Assessment Status': 42, 'LIVE DB - Planning Zone': 43, 'LIVE DB - PLM Overlay': 44, 'LIVE DB - Section': 45, 'LIVE DB - Township': 46, 'Related Admin Setup (=1)': 47, 'New/Existing': 49, 'Resolution: Approve/Reject': 50, 'Resolution: Reason': 51, 'Resolved? (1,0)': 52, 'Notify this user about Rejection': 53, 'Rejection notification text to insert in email message': 54, 'OK to Archive?': 55, 'OK to Archive? (1,0)': 56, 'Date / Time Archived': 57, 'Current Date Time': 58, 'Copy to Archive': 59, 'Warning Icon': 60, 'Import_Warning_Description____________________': 61, 'Warning?': 62, 'Warning? (1,0)': 63, 'Current User': 64, 'LIVE DB - Native Title Exists?': 65, '# of Native Titles': 66, '# of Warnings': 67, '# Resolved': 68, 'Record ID# of Entry to be Archived - Native Title': 69, 'ACT (Import)': 70, 'Area (Ha) (Import)': 71, 'Crown Allotment (Import)': 72, 'Parish (Import)': 73, 'Section (Import)': 74, 'Township (Import)': 75, 'Easting (Import)': 76, 'Forest Type (Import)': 77, 'Land Manager (Import)': 78, 'Land Status (Import)': 79, 'Length (km) (Import)': 80, 'Local Government Area (Import)': 81, 'Northing (Import)': 82, 'Native Title Name (Import)': 83, 'NT Assessment Status (Import)': 84, 'Native Title Status (Import)': 85, 'Overlay Code (Import)': 86, 'PLM Overlay (Import)': 87, 'Planning Zone (Import)': 88, 'Related Works Number (Import)': 89, 'Parcel Number (Import)': 90, 'Issue: ACT?': 91, 'Issue: Land Manager?': 93, 'Issue: Land Status?': 95, 'Issue: NT Assessment Status?': 97, 'Issue: Native Title Status?': 99, 'Issue: Overlay Code?': 101, 'Issue: ACT': 92, 'Issue: Land Manager': 94, 'Issue: Land Status': 96, 'Issue: NT Assessment Status': 98, 'Issue: Native Title Status': 100, 'Issue: Overlay Code': 102, 'Record ID# is the one to be Archived? Native Title': 103, 'Date Created': 1, 'Date Modified': 2, 'Last Modified By': 5, 'Record ID#': 3, 'Record Owner': 4}
# dictionary_map_WorkDetail = {'Warning Icon': 125, 'Import_Warning_Description____________________': 124, 'OBJECTID': 7, 'DAP_YEAR': 8, 'DAP_REF_NO': 9, 'DAP_NAME': 10, 'SCHEDULE': 11, 'RISK_LVL': 12, 'AREA_HA': 13, 'DISTRICT': 14, 'WORKCENTRE': 15, 'CREATOR': 17, 'CREATEDATE': 18, 'LASTEDITOR': 19, 'LASTDATE': 20, 'ORIG_FID': 21, 'OPS_PROG': 24, 'PLANT_2': 29, 'PLANT_3': 31, 'ASSET_ID': 34, 'ASSET_TYPE': 35, 'NATIVE_VEG': 37, 'PROPONENT': 38, 'REP_RISK': 39, 'NV_AREA_HA': 45, 'CB_AREA_HA': 51, 'Length_km': 56, 'YEAR_WORKS': 58, 'Shape_Area': 60, 'works program year (Import)': 95, 'first planned (Import)': 96, 'works number (Import)': 97, 'works name (Import)': 98, 'work schedule (Import)': 99, 'DAP/DMP risk level (Import)': 100, 'DISTRICT (Import)': 101, 'WORKCENTRE (Import)': 102, 'detailed works description (Import)': 103, 'priority for assessment (Import)': 104, 'operational program (Import)': 105, 'works category (Import)': 106, 'works detail (Import)': 107, 'possible plant (Import)': 108, 'did the district identify that they would follow best practive work methods (Import)': 109, 'best practice comments (Import)': 110, 'soil disturbance (Import)': 111, 'Start Of Diff SD': 186, 'Start Of Diff BP': 189, 'Start Of Diff CU': 190, 'Start Of Diff DWD': 191, 'Start Of Diff VC': 193, 'Start Of Diff JMP': 285, 'Start Of Diff SM': 289, 'Start Of Diff AB': 290, 'Changed Text SD': 187, 'Changed Text BP': 194, 'Changed Text CU': 195, 'Changed Text DWD': 196, 'Changed Text VC': 198, 'Changed Text JMP': 288, 'Changed Text SM': 291, 'Changed Text AB': 292, 'veg clearing (Import)': 112, 'PROPONENT (Import)': 113, 'reputational risk (Import)': 114, 'controversial (Import)': 115, 'wow required? (Import)': 116, 'chemical use (Import)': 117, 'activity status (Import)': 118, 'YEAR WORKS (Import)': 119, 'length (Import)': 120, 'EASTING (Import)': 121, 'northing (Import)': 122, 'area ha (Import)': 123, 'Date Created': 1, 'Date Modified': 2, 'Last Modified By': 5, 'Record ID#': 3, 'Record Owner': 4, 'Warning?': 128, 'Warning? (1,0)': 127, 'Related Admin Setup (=1)': 130, '# of Value Assessments': 132, '# of Warnings': 133, 'Related Works Number (formula)': 136, 'LIVE DB - Possible Plant': 138, 'LIVE DB - Detailed Works Description': 139, 'LIVE DB - Chemical Use': 140, 'LIVE DB - best practice comments': 141, 'LIVE DB - Veg Clearing': 142, 'LIVE DB - Soil Disturbance': 143, 'Issue: Best Practice': 144, 'Issue: Best Practice?': 150, 'Issue: Chemical Use': 145, 'Issue: Chemical Use?': 151, 'Issue: Detailed Works Description': 146, 'Issue: Detailed Works Description?': 152, 'Issue: Possible Plant': 147, 'Changed Text PP': 241, 'Issue: Possible Plant?': 153, 'Issue: Soil Disturbance': 148, 'Issue: Soil Disturbance?': 154, 'Issue: Veg Clearing': 149, 'Issue: Joint Managed Park': 286, 'Issue: Scripted Mitigation Summary': 293, 'Issue: Approval Body for Permits': 294, 'Issue: Veg Clearing?': 155, 'Issue: Joint Managed Park?': 287, 'Issue: Scripted Mitigation Summary?': 295, 'Issue: Approval Body for Permits?': 296, 'LIVE DB - Possible Plant (text)': 162, 'LIVE DB - Approval Body(s) for Permits (text)': 297, 'New/Existing': 164, 'LIVE DB - District': 172, 'Resolution: Approve/Reject': 173, 'Resolution: Reason': 174, 'Resolved? (1,0)': 177, '# Resolved': 178, 'LIVE DB - Value Assessment Exists?': 179, 'Current Date / Time': 180, 'Date / Time Archived': 181, 'Record ID# of Entry to be Archived - Value Assessments': 182, 'Record ID# is the one to be Archived? Value Assessment': 183, 'Copy to Archive': 184, 'Current User': 185, 'Text spacing': 199, 'PUBLISHDATE': 202, 'DESCRIPTION': 203, 'PRIORITY_WORKS': 204, 'WORKS_CATEGORY': 205, 'WORKS_TYPE_2': 207, 'WORKS_TYPE_3': 208, 'BEST_PRACTICE': 209, 'BEST_PRACTICE_COMMENTS': 210, 'SOIL_DISTURB': 211, 'CONTROVERSIAL': 212, 'WOW_REQUIRED': 213, 'CHEMICAL_USE': 214, 'ESTIMATED_COST': 215, 'FUND_SOURCE': 216, 'NVR_ASSET_STATUS': 217, 'NVR_CLEAR_TYPE': 220, 'NVR_Comments': 221, 'CB_Comments': 225, 'Activity_Status': 226, 'Shape_Length': 227, 'Notify this user about Rejection': 234, 'Rejection notification text to insert in email message': 235, 'OK to Archive?': 242, 'WORKS_TYPE_1': 243, 'PLANT_1': 244, 'NVR_ASSET': 245, 'NVR_LANDMANGEMENT': 246, 'CB_INC_EXT_SEC': 247, 'CB_IMPROVECONDITION': 248, 'CB_FUNDSOURCE': 249, 'Easting': 250, 'Northing': 251, 'OK to Archive? (1,0)': 252, 'CH_SENS': 254, 'FFMV Heritage Specialist Summary': 255, 'LAND_MANGR': 256, 'JointManagedPark': 257, 'PV_District': 258, 'CH_RAP': 259, 'Sites within 500m': 260, 'SCRIPT_DATE': 261, 'LIVE DB - Date of last ACHRIS check (Works Level)': 262, 'Date of last ACHRIS check (Works Level) (import)': 269, 'LIVE DB - Are there recorded Aboriginal Places within 500m': 263, 'Are there recorded Aboriginal Places within 500m (import)': 270, 'Approval Body(s) for Permits (multi-select) (import)': 281, 'PV District (import)': 282, 'LIVE DB - Approval Body(s) for Permits (multi-select)': 264, 'LIVE DB - PV District': 265, 'LIVE DB - Land Manager(old)': 266, 'Land Manager (import)': 273, 'LIVE DB - Scripted Mitigation Summary': 267, 'Scripted Mitigation Summary (import)': 274, 'LIVE DB - Sensitive Layer Intersect': 268, 'Sensitive Layer Intersect (import)': 275, 'LIVE DB - Joint Managed Park(old)': 283, 'Joint Managed Park (import)': 284, 'LIVE DB - Land Manager': 298, 'LIVE DB - Joint Managed Park': 299, 'LIVE DB - DAP / DMP risk level': 300, 'Issue: DAP / DMP risk level?': 302, 'Issue: DAP/DMP risk level': 304}
# dictionary_map_2 = {'New/Existing': 45, 'Works Number': 6, 'Works Name': 7, 'Updated Mitigation Advice From NEP 2021/2022': 24, 'Warning Icon': 61, 'Import_Warning_Description____________________': 56, 'Resolution: Approve/Reject': 53, 'Resolution: Reason': 54, 'Date / Time Archived': 68, '# of NEP Advice': 64, 'Biodiversity Risk Register Advice': 22, 'Copy to Archive': 69, 'DAP/DMP Risk Level': 8, 'Date Checked by NEP': 26, 'Date Checked by NEP (Import)': 40, 'Date Created': 1, 'Date Modified': 2, 'Detailed Works Description': 13, 'Does the works include chemical use?': 20, 'Is the value susceptaible to chemical use?': 21, 'Is the value susceptible to soil disturbance?': 15, 'Is the value susceptible to veg clearing?': 17, 'Issue: Date Checked by NEP': 46, 'Issue: Date Checked by NEP?': 47, 'Issue: NEP Reviewer': 49, 'Issue: NEP Reviewer?': 51, 'Issue: Updated Mitigation Advice From NEP 2021/2022': 50, 'RH Issue: Updated Mitigation Advice': 76, 'Issue: Updated Mitigation Advice From NEP 2021/2022?': 52, 'Last Modified By': 5, 'LIVE DB - Biodiversity Values Exists?': 48, 'LIVE DB - Date Checked by NEP': 35, 'LIVE DB - NEP reviewer': 36, 'LIVE DB - Updated Mitigation Advice From NEP 2021/2022': 37, 'NEP Mitigation to field staff': 23, 'NEP reviewer': 25, 'NEP reviewer (Import)': 39, 'NEW_KEY (Text)': 27, 'Notify this user about Rejection': 60, 'OK to Archive?': 66, 'OK to Archive? (1,0)': 67, 'Record ID#': 3, 'Record ID# is the one to be Archived? NEP Advice': 70, 'Record Owner': 4, 'Rejection notification text to insert in email message': 59, 'Related Admin Setup (=1)': 62, 'Reported soil disturbance category': 14, 'Reported Veg clearing category': 16, 'Resolved? (1,0)': 55, 'Updated Mitigation Advice From NEP 2021/2022 (Import)': 38, 'Value': 10, 'Value Description': 9, 'Value susceptible to waterway disturbance?': 19, 'Warning?': 57, 'Warning? (1,0)': 58, 'WOW permit required?': 18, 'X of Value': 11, 'Y of Value': 12, 'Current User': 71, 'Current Date/Time': 72, 'Record ID# of Entry to be Archived - NEP Advice': 73, 'TEST Key Formula': 74, 'LIVE DB - Updated Mitigation Advice': 75, 'LIVE DB - Standard Mitigation': 77, 'LIVE DB - NEP Mitigation to field staff': 78}
# dictionary_map_CH = {'QB_ID': 21, 'UNIQUE_ID': 7, 'ACHRIS_ID': 11, 'OBJECTID': 6, 'NAME': 8, 'DISTRICT': 9, 'DESCRIPTION': 10, 'ACHRIS_STATUS': 13, 'PLACE_TYPE': 14, 'PLACE_NAME': 15, 'BUFFER_TYPE': 18, 'FIRE_SENSITIVITY': 19, 'SCRIPT_DATE': 20, 'Works Number (Import)': 22, 'Works Name (Import)': 23, 'Detailed Works Description (Import)': 24, 'Place Number (Import)': 25, 'Place Name (Import)': 26, 'X coord of site (Import)': 27, 'Y coord of site (Import)': 28, 'Site Type (Import)': 29, 'ACHRIS Status (Import)': 30, 'Date Site Attached to Works (Import)': 31, 'District (Import)': 32, 'DEAD Issue: ACHRIS Date FIX???': 33, 'Issue: Date Site Attached to Works': 98, 'Issue: ACHRIS Date?': 34, 'Issue: Date Site Attached to Works?': 99, 'New/Existing': 35, 'Resolution: Approve/Reject': 36, 'Resolution: Reason': 37, 'Resolved? (1,0)': 38, 'Notify this user about Rejection': 39, 'Rejection notification text to insert in email message': 40, 'OK to Archive?': 41, 'OK to Archive? (1,0)': 42, 'Date / Time Archived': 43, 'Current Date Time': 44, 'Copy to Archive': 45, 'Warning Icon': 46, 'Import_Warning_Description____________________': 47, 'Warning?': 48, 'Warning? (1,0)': 49, 'Current User': 50, 'Related Admin Setup (=1)': 52, '# of Cultural Heritage Sites': 54, '# of Warnings': 55, '# of Resolved': 56, 'Record ID# of Entry to be Archived - Cultural Heritage Sites': 57, 'LIVE DB - Cultural Heritage Site Exists?': 58, 'LIVE DB - District': 59, 'LIVE DB - Date Site Attached to Works': 60, 'LIVE DB - ACHRIS Status': 61, 'LIVE DB - Site Type': 62, 'LIVE DB - Y coord of site': 63, 'LIVE DB - X coord of site': 64, 'LIVE DB - Place Name': 65, 'LIVE DB - Place Number': 66, 'LIVE DB - Detailed Works Description': 67, 'LIVE DB - Works Name': 68, 'LIVE DB - Related Works Number': 69, 'Record ID# is the one to be Archived - Cultural Heritage Sites': 70, 'LIVE DB - Date of last ACHRIS check (from Values Assessment)': 71, 'LIVE DB - Days since ACHRIS last checked for sites (from Values Assessment)': 72, 'LIVE DB - Were Additional Sites Identified? (from Values Assessment)': 73, 'LIVE DB - Has The Advice Been Updated By The Heritage Specialist? (from Values Assessment)': 74, 'Date of last ACHRIS check (from Values Assessment)': 77, 'ACHRIS_DATEMODIFIED': 78, 'CH_SENS': 79, 'FFMV Heritage Specialist Summary': 80, 'Site_X': 81, 'Site_Y': 82, 'RISK_LVL': 83, 'LAND_MANGR': 84, 'JointManagedPark': 85, 'CH_RAP': 86, 'Gurnaikurnai RAP': 87, 'Taungurung RAP': 88, 'Wurundjeri RAP': 89, 'No RAP': 90, 'Sites within 500m': 91, 'Bunurong RAP': 92, 'Works_Easting': 93, 'Works_Northing': 94, 'Works_prog_yr': 95, 'PV_District': 96, 'IS90DAY': 97, 'QB_ID (lookup)': 100, 'LIVE DB QB_ID - TEMP CUL VAL KEY (formula = Works#|Place#': 105, 'Date Created': 1, 'Date Modified': 2, 'Last Modified By': 5, 'Record ID#': 3, 'Record Owner': 4}
# dictionary_map_BIOFOR = {'QB_ID': 20, 'UNIQUE_ID': 7, 'COMM_NAME': 16, 'SCI_NAME': 15, 'MR_CODE': 14, 'X': 18, 'Y': 19, 'OBJECTID': 6, 'DISTRICT (raw)': 8, 'NAME': 9, 'DESCRIPTION': 10, 'RISK_LVL': 11, 'Mitigation': 12, 'EXTRA_INFO': 13, 'TYPE': 17, 'Soil_Disturbance': 22, 'Veg_Alteration': 23, 'Waterway_Disturbance': 24, 'Chemical_Use': 25, 'Last_Modified': 26, 'QB_ID test (formula)': 104, 'QB_ID import and formula match?': 105, 'Works Number (Import)': 54, 'District (Import)': 55, 'Works Name (Import)': 56, 'X of Value (Import)': 57, 'Y of Value (Import)': 58, 'Scripted Mitigation (Import)': 59, 'Chem Use (Import)': 60, 'Detailed Works Description (Import)': 61, 'Soil Disturb (Import)': 62, 'Veg Alteration (Import)': 63, 'Waterway Disturbance (Import)': 64, 'Value (Import)': 65, 'Value Type (Import)': 66, 'Warning Icon': 67, 'Import_Warning_Description____________________': 68, 'Warning?': 69, 'Warning? (1,0)': 70, 'Issue: Chem Use': 71, 'Issue: Chemical Use?': 72, 'Issue: Waterway Disturbance': 73, 'Issue: Waterway Disturbance?': 74, 'Issue: Scripted Mitigation': 75, 'Issue: Scripted Mitigation?': 76, 'Issue: Soil Disturb': 77, 'Issue: Soil Disturbance?': 78, 'Issue: Veg Alteration': 79, 'Issue: Veg Alteration?': 80, 'New/Existing': 81, 'Resolution: Approve/Reject': 82, 'Resolution: Reason': 83, 'Resolved? (1,0)': 84, 'Notify this user about Rejection': 85, 'Rejection notification text to insert in email message': 86, 'OK to Archive?': 87, 'OK to Archive? (1,0)': 88, 'Date / Time Archived': 89, 'Current Date Time': 90, 'Copy to Archive': 91, 'Record ID# is the one to be Archived? Biodiversity Values': 102, 'Current User': 103, 'Value Description (Import)': 106, 'Value Group (Import)': 141, 'Value_Group': 131, 'FM_Value': 132, 'FM_Description': 133, 'Script Date (Import)': 139, 'DATE_CHECKED': 134, '# of Biodiversity Values': 98, '# of Warnings': 99, '# Resolved': 100, 'Date Created': 1, 'Date Modified': 2, 'Last Modified By': 5, 'LIVE DB - Biodiversity Values Exists?': 107, 'LIVE DB - Chem Use': 125, 'LIVE DB - Detailed Works Description': 114, 'LIVE DB - District': 120, 'LIVE DB - Mitigation': 116, 'LIVE DB - Soil Disturb': 126, 'LIVE DB - Value': 110, 'LIVE DB - Value Description': 108, 'LIVE DB - Value Type': 109, 'LIVE DB - Veg Alteration': 112, 'LIVE DB - Waterway Disturbance': 111, 'LIVE DB - Works Name': 119, 'LIVE DB - Works Number': 121, 'LIVE DB - X of Value': 118, 'LIVE DB - Y of Value': 117, 'LIVE DB - Value Group': 138, 'Record ID#': 3, 'Record ID# of Entry to be Archived - Biodiversity Values': 101, 'Record Owner': 4, 'Related Admin Setup (=1)': 92, 'LIVE DB - Scripted Mitigation': 142, 'LIVE DB - Mitigation from NEP': 143, 'LIVE DB - NEP Specific Mitigation': 144, 'RECORD_ID': 145, 'STARTDATE': 146, 'MAX_ACC_KM': 147, 'COLLECTOR': 148, 'FFG_DESC': 149, 'EPBC_DESC': 150, 'TAXON_MOD': 151, 'VBA Initial Observation Date (Import)': 152, 'Bio GIS Record ID (Import)': 153, 'VBA Accuracy (km) (Import)': 154, 'VBA Collector (Import)': 155, 'FFG Act Status (Import)': 156, 'EPBC Act Status (Import)': 157, 'Taxon Last Modified Date (Import)': 158}
# def csv_to_quickbase_payload(csv_file_path, table_id, field_map):
#     payload = {
#         "to": table_id,
#         "data": []
#     }
#
#     with open(csv_file_path, mode='r', newline='') as csvfile:
#         reader = csv.DictReader(csvfile)
#
#         for row in reader:
#             record = {}
#             for csv_column, quickbase_id in field_map.items():
#                 if csv_column in row:  # Check if the column exists in the row
#                     # Create the Quickbase field dictionary with "value"
#                     record[str(quickbase_id)] = {"value": row[csv_column]}
#
#             payload["data"].append(record)
#
#     # Set "fieldsToReturn" to only include the field IDs in "data"
#     payload["fieldsToReturn"] = list(field_map.values())
#     return payload
#
# # convert .xls to .csv here
# for xls in glob.iglob(os.path.join(workpath, '*.xlsx')):
#     read_file = pd.read_excel(xls)
#     read_file.to_csv(xls[:-5] + '.csv', index=None, header=True, encoding='utf-8')
#     header_list = read_file.columns.tolist()
#     # Determine the correct dictionary map
#     if mode != "JFMP" or mode != "NBFT":
#         if 'ACHRIS_ID' in header_list and 'ACHRIS_STATUS' in header_list:
#             selected_map = dictionary_map_CH
#             table_id = 'bunqzgv3m'  # CH ScratchPad
#             print("Using dictionary_map_CH")
#         # elif "SCI_NAME" in header_list and "FFG_DESC" in header_list: # to be developed later
#         #     selected_map = dictionary_map_2
#         #     table_id = 'br3nvvq65'  # NEP ScratchPad
#         #     print("Using dictionary_map_2")
#         elif "TAXON_MOD" in header_list and "FM_Description" in header_list:
#             selected_map = dictionary_map_BIOFOR
#             table_id = 'brhc4qxmb'  # Bio & Forest Values ScratchPad
#             print("Using dictionary_map_BIOFOR")
#         elif "C_PARISH" in header_list and "C_SEC" in header_list:
#             selected_map = dictionary_map_NT
#             table_id = 'bt7r5uggj'  # Native Title ScratchPad
#             print("Using dictionary_map_NT")
#         elif "ASSET_ID" in header_list and "YEAR_WORKS" in header_list:
#             selected_map = dictionary_map_WorkDetail
#             table_id = 'bres8ggfp'  # Values Assessments ScratchPad
#             print("Using dictionary_map_WorkDetail")
#         else:
#             selected_map = None
#             print("No matching dictionary map found!")
#
#         # Output the selected map
#         if selected_map:
#             print(f"Selected Map: {selected_map}")
#             csv_file_path = xls[:-5] + '.csv'
#             payload = csv_to_quickbase_payload(csv_file_path, table_id, selected_map)
#         else:
#             print("No dictionary map selected. Unable to create payload.")
#         # Usage example
#
#         if QuickBaseAPI == True:
#             ###REST API Connection to Quickbase using Requests
#             headers = {
#                 'Content-Type': 'application/json',
#                 'QB-Realm-Hostname': 'mlutze.quickbase.com',
#                 'User-Agent': 'Python_UpSert',
#                 'Authorization': 'QB-USER-TOKEN b9emt9_cy4u_0_bfxnakh97q8njcdcenvkcsc8twi'
#             }
#             body = {}
#             r = requests.post('https://api.quickbase.com/v1/records', headers=headers, json=payload)
#             print(r.ok)
#             print(json.dumps(r.json(), indent=4))
#     os.remove(xls)
print("Successfully finished")

ended_time = datetime.datetime.now()
duration = ended_time - started_time
features_processed = arcpy.management.GetCount(worksfc)
print(f"Processing time: Start: {start_time}   Finish: {ended_time.strftime('%H:%M:%S')}")
print(f"Time taken to process: {duration}")
print(f"Features processed: {features_processed}")
print(f"Script reference: {start_date}{district_name}{mode}")
