import pandas as pd
import pyodbc as py
import shapely
from shapely.wkt import loads
from shapely.geometry import mapping, Polygon
import fiona
import numpy as np
import os
import sys
import paramiko
import configparser

def setup(scenario_name, folder_path):
    #Make a folder in folder_path for scenario_name
    if not os.path.exists(folder_path + '\\' + scenario_name):
        os.makedirs(folder_path + '\\' + scenario_name)
    #Connect to Hazus SQL Server database where scenario results are stored
    comp_name = os.environ['COMPUTERNAME']
    cnxn = py.connect('Driver=ODBC Driver 11 for SQL Server;SERVER=' + comp_name + '\HAZUSPLUSSRVR;DATABASE=' + scenario_name + ';UID=hazuspuser;PWD=Gohazusplus_02')
    return comp_name, cnxn

def read_sql(comp_name, cnxn, scenario_name):
    #Select Hazus results from SQL Server scenario database
    sql_econ_loss = """SELECT Tract, SUM(ISNULL(BldgLoss, 0) + ISNULL(ContentLoss, 0) + ISNULL(InvLoss, 0) + ISNULL(RelocLoss, 0) +
    ISNULL(IncLoss, 0) + ISNULL(RentLoss, 0) + ISNULL(WageLoss, 0)) AS EconLoss
    FROM %s.dbo.[eqTractEconLoss] GROUP BY [eqTractEconLoss].Tract""" %scenario_name
    sql_county_fips = """SELECT Tract, CountyFips FROM %s.dbo.[hzTract]""" %scenario_name
    sql_demographics = """SELECT Tract, Population, Households FROM
    %s.dbo.[hzDemographicsT]""" %scenario_name
    sql_impact = """SELECT Tract, DebrisW, DebrisS, DisplacedHouseholds AS DisplHouse, ShortTermShelter AS Shelter
    FROM %s.dbo.[eqTract]""" %scenario_name
    sql_injury = """SELECT Tract, SUM(ISNULL(Level1Injury, 0)) AS Level1Injury, SUM(ISNULL(Level2Injury, 0)) AS Level2Injury,
    SUM(ISNULL(Level3Injury, 0)) AS Level3Injury, SUM(ISNULL(Level1Injury, 0) + ISNULL(Level2Injury, 0) + ISNULL(Level3Injury, 0)) AS NonFatal5p
    FROM %s.dbo.[eqTractCasOccup] WHERE CasTime = 'C' AND InOutTot = 'Tot' GROUP BY Tract""" %scenario_name
    sql_building_damage = """SELECT Tract, SUM(ISNULL(PDsNoneBC, 0)) As NoDamage, SUM(ISNULL(PDsSlightBC, 0) + ISNULL(PDsModerateBC, 0)) AS GreenTag,
    SUM(ISNULL(PDsExtensiveBC, 0)) AS YellowTag, SUM(ISNULL(PDsCompleteBC, 0)) AS RedTag FROM %s.dbo.[eqTractDmg] WHERE DmgMechType = 'STR' GROUP BY Tract""" %scenario_name
    sql_building_damage_occup = """SELECT Occupancy, SUM(ISNULL(PDsNoneBC, 0)) As NoDamage, SUM(ISNULL(PDsSlightBC, 0) + ISNULL(PDsModerateBC, 0)) AS GreenTag,
    SUM(ISNULL(PDsExtensiveBC, 0)) AS YellowTag, SUM(ISNULL(PDsCompleteBC, 0)) AS RedTag FROM %s.dbo.[eqTractDmg] WHERE DmgMechType = 'STR' GROUP BY Occupancy""" %scenario_name
    sql_building_damage_bldg_type = """SELECT eqBldgType, SUM(ISNULL(PDsNoneBC, 0)) As NoDamage, SUM(ISNULL(PDsSlightBC, 0) + ISNULL(PDsModerateBC, 0)) AS GreenTag,
    SUM(ISNULL(PDsExtensiveBC, 0)) AS YellowTag, SUM(ISNULL(PDsCompleteBC, 0)) AS RedTag FROM %s.dbo.[eqTractDmg] WHERE DmgMechType = 'STR' GROUP BY eqBldgType""" %scenario_name
    sql_tract_spatial = """SELECT Tract, Shape.STAsText() AS Shape FROM %s.dbo.[hzTract]""" %scenario_name
    #Group tables and queries into iterable
    hazus_results = {'econ_loss': sql_econ_loss , 'county_fips': sql_county_fips, 'demographics': sql_demographics, 'impact': sql_impact, 'injury': sql_injury,
                    'building_damage': sql_building_damage, 'building_damage_occup': sql_building_damage_occup, 'building_damage_bldg_type': sql_building_damage_bldg_type,
                    'tract_spatial': sql_tract_spatial}
    #Read Hazus result tables from SQL Server into dataframes with Tract ID as index
    hazus_results_df = {table: pd.read_sql(query, cnxn) for table, query in hazus_results.items()}
    for name, dataframe in hazus_results_df.items():
        if (name != 'building_damage_occup') and (name != 'building_damage_bldg_type'):
            dataframe.set_index('Tract', append=False, inplace=True)
    return hazus_results_df

def results(hazus_results_df):
        #Join and group Hazus result dataframes into final TwoPAGER dataframes
        tract_results = hazus_results_df['county_fips'].join([hazus_results_df['econ_loss'], hazus_results_df['demographics'], hazus_results_df['impact'],
                        hazus_results_df['injury'], hazus_results_df['building_damage']])
        county_results = tract_results.groupby('CountyFips').sum()
        return tract_results, county_results

def to_csv(hazus_results_df, tract_results, county_results, folder_path, scenario_name):
        #Export TwoPAGER dataframes to text files
        two_pager_df = {'building_damage_occup': hazus_results_df['building_damage_occup'], 'building_damage_bldg_type': hazus_results_df['building_damage_bldg_type'],
                            'tract_results': tract_results, 'county_results': county_results}
        for name, dataframe in two_pager_df.items():
            path = folder_path + '\\' + scenario_name + '\\' + name + '.txt'
            dataframe.to_csv(path)

def to_shp(folder_path, scenario_name, hazus_results_df, tract_results):
        #Create shapefile of TwoPAGER tract table
        schema = {
            'geometry': 'Polygon',
            'properties': {'Tract': 'str',
                            'CountyFips': 'int',
                            'EconLoss': 'float',
                            'Population': 'int',
                            'Households': 'int',
                            'DebrisW': 'float',
                            'DebrisS': 'float',
                            'DisplHouse': 'float',
                            'Shelter': 'float',
                            'NonFatal5p': 'float',
                            'NoDamage': 'float',
                            'GreenTag': 'float',
                            'YellowTag': 'float',
                            'RedTag': 'float'},
                            }
        with fiona.open(path=folder_path + '\\' + scenario_name + '\\tract_results.shp', mode='w', driver='ESRI Shapefile', schema=schema, crs={'proj':'longlat', 'ellps':'WGS84', 'datum':'WGS84', 'no_defs':True}) as c:
            for index, row in hazus_results_df['tract_spatial'].iterrows():
                for i in row:
                    tract = shapely.wkt.loads(i)
                    c.write({'geometry': mapping(tract),
                            'properties': {'Tract': index,
                                            'CountyFips': str(np.int64(tract_results.loc[str(index), 'CountyFips'])),
                                            'EconLoss': str(np.float64(tract_results.loc[str(index), 'EconLoss'])),
                                            'Population': str(np.int64(tract_results.loc[str(index), 'Population'])),
                                            'Households': str(np.int64(tract_results.loc[str(index), 'Households'])),
                                            'DebrisW': str(np.float64(tract_results.loc[str(index), 'DebrisW'])),
                                            'DebrisS': str(np.float64(tract_results.loc[str(index), 'DebrisS'])),
                                            'DisplHouse': str(np.float64(tract_results.loc[str(index), 'DisplHouse'])),
                                            'Shelter': str(np.float64(tract_results.loc[str(index), 'Shelter'])),
                                            'NonFatal5p': str(np.float64(tract_results.loc[str(index), 'NonFatal5p'])),
                                            'NoDamage': str(np.float64(tract_results.loc[str(index), 'NoDamage'])),
                                            'GreenTag': str(np.float64(tract_results.loc[str(index), 'GreenTag'])),
                                            'YellowTag': str(np.float64(tract_results.loc[str(index), 'YellowTag'])),
                                            'RedTag': str(np.float64(tract_results.loc[str(index), 'RedTag']))}
                                            })
        c.close()

def to_ftp(folder_path, scenario_name):
    #Upload TwoPAGER text files and shapefile to FEMA FTP site
    load_files = ['building_damage_occup.txt', 'building_damage_bldg_type.txt', 'tract_results.txt',
                    'county_results.txt', 'tract_results.dbf', 'tract_results.shp']
    credentials = {'host': 'data.femadata.com', 'username': 'jordan.burns', 'password': 'Tt&jrayEZYL[*q2+',
                    'dir': '/FIMA/NHRAP/Earthquake/TwoPAGER'}
    transport = paramiko.Transport(credentials['host'], 22)
    transport.connect(username=credentials['username'], password=credentials['password'])
    sftp = paramiko.SFTPClient.from_transport(transport)
    sftp.chdir(credentials['dir'])
    dirs = sftp.listdir()
    if str(scenario_name) not in dirs:
        sftp.mkdir(scenario_name)
        sftp.chmod((credentials['dir'] + '/' + scenario_name), mode=777)
    for i in load_files:
        local_path = folder_path + '\\' + scenario_name + '\\' + i
        remote_path = './' + scenario_name + '//' + i
        sftp.put(local_path, remote_path)
        sftp.chmod(remote_path, mode=777)
    sftp.close()
    return credentials

#Roll up subfunctions into one overall function
def two_pager(scenario_name, folder_path, ftp=False):
    comp_name, cnxn = setup(scenario_name, folder_path)
    hazus_results_df = read_sql(comp_name, cnxn, scenario_name)
    tract_results, county_results = results(hazus_results_df)
    to_csv(hazus_results_df, tract_results, county_results, folder_path, scenario_name)
    to_shp(folder_path, scenario_name, hazus_results_df, tract_results)
    if ftp == True:
        credentials = to_ftp(folder_path, scenario_name)
        print('Hazus earthquake results available for download at: https://' + credentials['host'] + credentials['dir'] + '/' + scenario_name)
    print('Hazus earthquake results available locally at: ' + folder_path + '\\' + scenario_name)

#Execute TwoPAGER function using variables given in Config.ini
# os.chdir('C:\code')
config = configparser.ConfigParser()
config.read('./Config.ini')
two_pager(config.get('VARIABLES', 'scenario_name'), config.get('VARIABLES', 'folder_path'), config.getboolean('VARIABLES', 'ftp'))
