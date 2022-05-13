import json
import pandas as pd
from itertools import chain
import quip
from flask import Flask, request, render_template, send_file
from flask_cors import CORS
from io import BytesIO

app = Flask(__name__)
cors = CORS(app)
app.config['CORS_HEADERS'] = 'Content-Type'


@app.route('/getassociate')
def getInputfromUser():

    global finalData, data, poolData, dataEx, programName, quantityReq, ipm_type, initialProgramName, initial, counter, selected, lastProgram, client, different_program, poolValue, jobTool, jobType, loop_check, programType, jsonData, poolName
    client = quip.QuipClient(
        access_token=
        "VmRGOU1BcFdxMFI=|1658339780|vfM+5aF19nY705P1u43PcosQYsLBSCsPujTPxLKjtnQ="
    )
    thread_id = client.get_thread('f25rA06XRJEP')
    dfs = pd.read_html(thread_id['html'])
    data = dfs[0]
    data.columns = data.iloc[0]
    data = data[1:]
    initial = False
    counter = 0
    #file = 'IPM-New-4.xlsx'
    #data = pd.read_excel(file, sheet_name='IPM Sheet')
    data.sort_values(by='Base Program', inplace=True)
    programName = ''
    ipm_type = ''
    quantityReq = ''
    initialProgramName = ''
    different_program = False
    poolValue = []
    selected = []
    finalData = []
    jobType = ""
    jobTool = ""
    programType = ""
    loop_check = 0
    jsonData = ''
    lastProgram = False
    dataEx = {
        "Associate ID": '',
        "Manager_ID": '',
        "Tenure (days)": '',
        "Shift schedule": '',
        "Base Program": '',
        "IPM Program": '',
        "Scorecard Rank": '',
        'Pentile': '',
        "IPM Difference": '',
        "Training Requirement": '',
        'IPM Type': []
    }
    poolData = {
        "HMI-Classification": ["NIKE", "JPOD", "Proxemics(Tron)", "SPA"],
        "HMI-Segementation": [
            "Robin-SLF", "Robin SFD", "Robin DI", "Shipshape", "TruckQueue",
            "Pallet Buffer"
        ],
        "SGMT-Classification": ["CPEX SIOC", "CPEX SIOB", "CPEX Polybag"],
        "SGMT-Segmentation":
        ["Canvas Lite tile", "Canvas Grey Scale Annotations"],
        "VBI Console-Segmentation": ["VBI", "Proxemics"],
        "AVANT-Classification":
        ["PVP", "Package Zero", "Transhipment", "Vapour"]
    }
    poolName = [
        "HMI-Classification", "HMI-Segementation", "SGMT-Classification",
        "SGMT-Segmentation", "VBI Console-Segmentation", "AVANT-Classification"
    ]
    jobType = {
        "HMI-Classification": "Video",
        "HMI-Segementation": "Video",
        "SGMT-Classification": "Video",
        "SGMT-Segmentation": "Image",
        "VBI Console-Segmentation": "Image",
        "AVANT-Classification": "Image"
    }
    programName = request.args.get('programName')
    quantityReq = request.args.get('quantityReq')
    ipm_type = request.args.get('ipm_type')
    initialProgramName = programName
    initial = True
    counter = 0
    selected = []
    lastProgram = False
    sorter()
    jsonResponse = printer()
    return jsonResponse


@app.route('/')
def home():
    return render_template('index.html')


def getJobTool():
    global poolData, programName, jobType
    for key, value in poolData.items():
        for i in value:
            if i == programName:
                for j, k in jobType.items():
                    if key == j:
                        return k


def getPoolRequired():
    global poolData, programName
    for key, value in poolData.items():
        for i in value:
            if i == programName:
                return value


def getNextInSamePool(poolValue):
    global different_program, lastProgram, poolName, poolData, lastProgram
    different_program = True

    i = poolValue.index(programName)
    program = getProgramType()
    p = poolName.index(program)
    if p == len(poolName) - 1:
        lastProgram = True
    if i < (len(poolValue)) - 1:
        return poolValue[i + 1]
    elif i == (len(poolValue)) - 1:
        for key, value in poolData.items():
            pool = getNext(poolName)
            if key == pool and pool != program:
                return value[0]


def getNext(valueList):
    for index, elem in enumerate(valueList):
        if (index + 1 < len(valueList) and index - 1 >= 0):
            next_el = str(valueList[index + 1])
    return next_el


def getProgramType():
    global poolData, programName
    for key, value in poolData.items():
        for i in value:
            if i == programName:
                return key


def initializer():
    global counter, programName, poolValue, jobTool, programType
    poolValue = getPoolRequired()
    jobTool = getJobTool()
    programType = getProgramType()


def sorter():
    initializer()
    while counter != int(quantityReq):
        getAssosicate()


def printer():
    global initial, jsonData, finalData
    headCount = {'comment': 'Head Count Short'}
    if initial == False:
        return 'Fatal Error'
    else:
        global selected
        if lastProgram == True:
            print("\n------<<<<<<  Head Count Short >>>>>-------\n")
        print("Following Associates are selected: \n")
        j = 0
        for i in selected:
            j += 1
            print(
                str(j) + ". " + i[0] + " is selected from " + i[1] + " " +
                i[2])

        for k in selected:
            dataEx = {}
            dataEx["Associate ID"] = k[0]
            dataEx["Manager_ID"] = k[3]
            dataEx["Tenure (days)"] = (k[7])
            dataEx["Shift schedule"] = (k[4])
            dataEx["Base Program"] = (k[1])
            dataEx["IPM Program"] = (initialProgramName)
            dataEx["Scorecard Rank"] = (k[5])
            dataEx['Pentile'] = (k[6])
            dataEx['IPM Difference'] = (k[8])
            dataEx['Training Requirement'] = (k[2])
            dataEx['IPM Type'] = (k[9])

            finalData.append(dataEx)
        if lastProgram == True:
            finalData.append(headCount)
        jsonData = json.dumps(finalData)
        initial = False
        return jsonData


@app.route('/download')
def downloadExcel():
    global finalData
    df = pd.DataFrame(finalData)
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Candidates', index=False)
    writer.save()
    output.seek(0)
    print("\n  Excel generated at IPM-Selected.xlsx")

    return send_file(output,
                     attachment_filename='IPM-Selected.xlsx',
                     as_attachment=True)


#excel sheet fields
def getAssosicate():
    global counter, loop_check, programName, initialProgramName, quantityReq, selected, different_program, data

    loop_check += 1
    for i, row in data.iterrows():
        associate_id = row['Associate ID']
        base_type = row['Base program type']
        base_program_tool = row['Base program job type']
        tenure_base_program = row['Tenure (days)']
        is_training = row['is in Training/ Nesting/Ramp-up/CDP/CFT']
        base_program = row['Base Program']
        scorecard_pentile = row['Scorecard Pentile']
        diff_ipm_program = row['IPM Difference']
        manager_id = row['Manager_ID']
        shift_schedule = row['Shift schedule']
        scorecard_rank = row['Scorecard Rank']
        if is_training == "Yes" or scorecard_pentile == "P5" or scorecard_pentile == 'P5-Bottom 10' or scorecard_pentile == 'Not eligible':
            continue
        elif associate_id == '\u200b' or base_type == '\u200b' or base_program_tool == '\u200b' or tenure_base_program == '\u200b' or base_program == '\u200b' or scorecard_pentile == '\u200b' or diff_ipm_program == '\u200b' or manager_id == '\u200b' or shift_schedule == '\u200b' or scorecard_rank == '\u200b':
            continue
        elif base_type == programType or base_program_tool == jobTool:

            if loop_check == 1:

                if int(diff_ipm_program) < 7 and int(tenure_base_program) > (
                        25 * 7
                ) and int(
                        diff_ipm_program
                ) != 0 and base_program != programName and base_program != initialProgramName:

                    if associate_id not in chain(*selected):
                        selected.append([
                            associate_id, base_program,
                            " Different Program - Training needs to be done"
                            if different_program == True else " ", manager_id,
                            shift_schedule, scorecard_rank, scorecard_pentile,
                            tenure_base_program, diff_ipm_program, ipm_type
                        ])
                        counter += 1
                        if counter == int(quantityReq):
                            break
            elif loop_check == 2:
                if int(diff_ipm_program) >= 7 and int(tenure_base_program) > (
                        25 * 7
                ) and base_program != programName and base_program != initialProgramName:
                    if associate_id not in chain(*selected):

                        selected.append([
                            associate_id, base_program,
                            " Different Program - Training needs to be done"
                            if different_program == True else
                            " Training to be provided", manager_id,
                            shift_schedule, scorecard_rank, scorecard_pentile,
                            tenure_base_program, diff_ipm_program, ipm_type
                        ])

                        counter += 1
                        if counter == int(quantityReq):
                            break
            elif loop_check == 3:
                if int(tenure_base_program) > (25 * 7) and int(
                        diff_ipm_program
                ) == 0 and base_program != programName and base_program != initialProgramName:
                    if associate_id not in chain(*selected):
                        selected.append([
                            associate_id, base_program,
                            " Different Program - Training needs to be done"
                            if different_program == True else
                            " Fresh Candidate - Training to be provided",
                            manager_id, shift_schedule, scorecard_rank,
                            scorecard_pentile, tenure_base_program,
                            diff_ipm_program, ipm_type
                        ])
                        counter += 1
                        if counter == int(quantityReq):
                            break
            elif loop_check == 4:

                programName = getNextInSamePool(poolValue)
                if lastProgram == True:

                    counter = int(quantityReq)
                    break
                else:
                    loop_check = 0
                    sorter()


def main():
    app.run(host="localhost", port=4000, debug=False)


if __name__ == "__main__":
    main()
