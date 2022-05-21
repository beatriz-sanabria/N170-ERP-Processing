from utils_ERP import init_eeg
from epochs import epochs
from signal_filtering import signal_filtering_tensor
from decomposition import parafac_decomposition, components, select_components
from print_erps_data import data_per_hemisphere_6, print_peak_measures, data_per_hemisphere_4
import xlsxwriter

routes = 

finalraw = {}
isGroupAnalysis = False

for route in routes:
    eeg = route

    events_eeg, raw, name, laterality, group, age, state, file_name, duration_bad_marks = init_eeg(eeg)
    # Without SSA
    raw2 = signal_filtering_tensor(raw)
    raw.drop_channels(['EKG', 'EOG'])
    epochsFaces, epochsObjects, times = epochs(events_eeg, raw)  # Separate raw in epochs per faces and objects
    epochsFaces_WSSA, epochsObjects_WSSA, times_WSSA = epochs(events_eeg,
                                                              raw2)  # Separate raw in epochs per faces and objects

    rank = 4

    # Faces
    title_Parafac = 'Faces \n (PARAFAC decomposition)'
    parafac_comp_faces, parafac_spatial_faces, comp_selected_faces, parafac_comp_WSSA_faces, \
    parafac_spatial_WSSA_faces, comp_selected_WSSA_faces = parafac_decomposition(epochsFaces, rank, times,
                                                                                 title_Parafac, epochsFaces_WSSA,
                                                                                 times_WSSA)

    # Objects
    title_Parafac = 'Objects \n (PARAFAC decomposition)'
    parafac_comp_objects, parafac_spatial_objects, comp_selected_objects, parafac_comp_WSSA_objects, \
    parafac_spatial_WSSA_objects, comp_selected_WSSA_objects = parafac_decomposition(epochsObjects, rank, times,
                                                                                     title_Parafac, epochsObjects_WSSA,
                                                                                     times_WSSA)

    faces, objects, factor_face, factor_obj = components(parafac_comp_faces, parafac_spatial_faces,
                                                         parafac_comp_objects, parafac_spatial_objects, times,
                                                         comp_selected_faces, comp_selected_objects)
    select_components(times, factor_face, factor_obj)

    faces_WSSA, objects_WSSA, factor_face_WSSA, factor_obj_WSSA = components(parafac_comp_WSSA_faces,
                                                                             parafac_spatial_WSSA_faces,
                                                                             parafac_comp_WSSA_objects,
                                                                             parafac_spatial_WSSA_objects, times,
                                                                             comp_selected_WSSA_faces,
                                                                             comp_selected_WSSA_objects)
    select_components(times, factor_face_WSSA, factor_obj_WSSA)

    print_peak_measures(faces, title='Faces')

    f_l_right_P1, f_l_left_P1, f_l_right_N170, f_l_left_N170, f_a_right_P1, f_a_left_P1, f_a_right_N170, f_a_left_N170 = \
        data_per_hemisphere_6(faces, title='Faces')

    f_l_right_P1_4, f_l_left_P1_4, f_l_right_N170_4, f_l_left_N170_4, f_a_right_P1_4, f_a_left_P1_4, f_a_right_N170_4, \
    f_a_left_N170_4 = data_per_hemisphere_4(faces, title='Faces')

    print_peak_measures(objects, title='Objects')

    o_l_right_P1, o_l_left_P1, o_l_right_N170, o_l_left_N170, o_a_right_P1, o_a_left_P1, o_a_right_N170, o_a_left_N170 = \
        data_per_hemisphere_6(objects, title='Objects')

    o_l_right_P1_4, o_l_left_P1_4, o_l_right_N170_4, o_l_left_N170_4, o_a_right_P1_4, o_a_left_P1_4, o_a_right_N170_4, \
    o_a_left_N170_4 = data_per_hemisphere_4(objects, title='Objects')

    workbook = xlsxwriter.Workbook(file_name)
    worksheet = workbook.add_worksheet("ERPs SSA")

    data = (['Name', name], ['Laterality', laterality], ['Group', group], ['Age', age], ['State', state], ['', ''],
            ['O1 Faces:', ''],
            ['P1 amplitude (µV)', faces[4]["P1 amplitude"]], ['P1 latency (ms)', faces[4]["P1 latency"]],
            ['N170 amplitude (µV)', faces[4]["N170 amplitude"]], ['N170 latency (ms)', faces[4]["N170 latency"]],
            ['', ''],
            ['O2 Faces:', ''],
            ['P1 amplitude (µV)', faces[5]["P1 amplitude"]], ['P1 latency (ms)', faces[5]["P1 latency"]],
            ['N170 amplitude (µV)', faces[5]["N170 amplitude"]], ['N170 latency (ms)', faces[5]["N170 latency"]],
            ['', ''],
            ['T5 Faces:', ''],
            ['P1 amplitude (µV)', faces[2]["P1 amplitude"]], ['P1 latency (ms)', faces[2]["P1 latency"]],
            ['N170 amplitude (µV)', faces[2]["N170 amplitude"]], ['N170 latency (ms)', faces[2]["N170 latency"]],
            ['', ''],
            ['T6 Faces:', ''],
            ['P1 amplitude (µV)', faces[3]["P1 amplitude"]], ['P1 latency (ms)', faces[3]["P1 latency"]],
            ['N170 amplitude (µV)', faces[3]["N170 amplitude"]], ['N170 latency (ms)', faces[3]["N170 latency"]],
            ['', ''],
            ['P3 Faces:', ''],
            ['P1 amplitude (µV)', faces[0]["P1 amplitude"]], ['P1 latency (ms)', faces[0]["P1 latency"]],
            ['N170 amplitude (µV)', faces[0]["N170 amplitude"]], ['N170 latency (ms)', faces[0]["N170 latency"]],
            ['', ''],
            ['P4 Faces:', ''],
            ['P1 amplitude (µV)', faces[1]["P1 amplitude"]], ['P1 latency (ms)', faces[1]["P1 latency"]],
            ['N170 amplitude (µV)', faces[1]["N170 amplitude"]], ['N170 latency (ms)', faces[1]["N170 latency"]],
            ['', ''],
            ['Faces Mean Latency per Hemisphere (6 electrodes):', ''],
            ['P1  left (ms)', f_l_left_P1], ['P1 right (ms)', f_l_right_P1],
            ['N170 left (ms)', f_l_left_N170], ['N170 right (ms)', f_l_right_N170],
            ['', ''],
            ['Faces Mean Amplitude per Hemisphere (6 electrodes):', ''],
            ['P1  left (µV)', f_a_left_P1], ['P1 right (µV)', f_a_right_P1],
            ['N170 left (µV)', f_a_left_N170], ['N170 right (µV)', f_a_right_N170],
            ['', ''], ['', ''],
            ['Faces Mean Latency per Hemisphere (4 electrodes):', ''],
            ['P1  left (ms)', f_l_left_P1_4], ['P1 right (ms)', f_l_right_P1_4],
            ['N170 left (ms)', f_l_left_N170_4], ['N170 right (ms)', f_l_right_N170_4],
            ['', ''],
            ['Faces Mean Amplitude per Hemisphere (4 electrodes):', ''],
            ['P1  left (µV)', f_a_left_P1_4], ['P1 right (µV)', f_a_right_P1_4],
            ['N170 left (µV)', f_a_left_N170_4], ['N170 right (µV)', f_a_right_N170_4],
            ['', ''], ['', ''],
            ['O1 Objects:', ''],
            ['P1 amplitude (µV)', objects[4]["P1 amplitude"]], ['P1 latency (ms)', objects[4]["P1 latency"]],
            ['N170 amplitude (µV)', objects[4]["N170 amplitude"]], ['N170 latency (ms)', objects[4]["N170 latency"]],
            ['', ''],
            ['O2 Objects:', ''],
            ['P1 amplitude (µV)', objects[5]["P1 amplitude"]], ['P1 latency (ms)', objects[5]["P1 latency"]],
            ['N170 amplitude (µV)', objects[5]["N170 amplitude"]], ['N170 latency (ms)', objects[5]["N170 latency"]],
            ['', ''],
            ['T5 Objects:', ''],
            ['P1 amplitude (µV)', objects[2]["P1 amplitude"]], ['P1 latency (ms)', objects[2]["P1 latency"]],
            ['N170 amplitude (µV)', objects[2]["N170 amplitude"]], ['N170 latency (ms)', objects[2]["N170 latency"]],
            ['', ''],
            ['T6 Objects:', ''],
            ['P1 amplitude (µV)', objects[3]["P1 amplitude"]], ['P1 latency (ms)', objects[3]["P1 latency"]],
            ['N170 amplitude (µV)', objects[3]["N170 amplitude"]], ['N170 latency (ms)', objects[3]["N170 latency"]],
            ['', ''],
            ['P3 Objects:', ''],
            ['P1 amplitude (µV)', objects[0]["P1 amplitude"]], ['P1 latency (ms)', objects[0]["P1 latency"]],
            ['N170 amplitude (µV)', objects[0]["N170 amplitude"]], ['N170 latency (ms)', objects[0]["N170 latency"]],
            ['', ''],
            ['P4 Objects:', ''],
            ['P1 amplitude (µV)', objects[1]["P1 amplitude"]], ['P1 latency (ms)', objects[1]["P1 latency"]],
            ['N170 amplitude (µV)', objects[1]["N170 amplitude"]], ['N170 latency (ms)', objects[1]["N170 latency"]],
            ['', ''],
            ['Objects Mean Latency per Hemisphere (6 electrodes):', ''],
            ['P1  left (ms)', o_l_left_P1], ['P1 right (ms)', o_l_right_P1],
            ['N170 left (ms)', o_l_left_N170], ['N170 right (ms)', o_l_right_N170],
            ['', ''],
            ['Objects Mean Amplitude per Hemisphere (6 electrodes):', ''],
            ['P1  left (µV)', o_a_left_P1], ['P1 right (µV)', o_a_right_P1],
            ['N170 left (µV)', o_a_left_N170], ['N170 right (µV)', o_a_right_N170],
            ['Objects Mean Latency per Hemisphere (4 electrodes):', ''],
            ['P1  left (ms)', o_l_left_P1_4], ['P1 right (ms)', o_l_right_P1_4],
            ['N170 left (ms)', o_l_left_N170_4], ['N170 right (ms)', o_l_right_N170_4],
            ['', ''],
            ['Objects Mean Amplitude per Hemisphere (4 electrodes):', ''],
            ['P1  left (µV)', o_a_left_P1_4], ['P1 right (µV)', o_a_right_P1_4],
            ['N170 left (µV)', o_a_left_N170_4], ['N170 right (µV)', o_a_right_N170_4])

    # Start from the first cell. Rows and columns are zero indexed.
    row = 0
    col = 0

    # Iterate over the data and write it out row by row.
    for name, score in data:
        worksheet.write(row, col, name)
        worksheet.write(row, col + 1, score)
        row += 1

    ##

    worksheet2 = workbook.add_worksheet("Faces SSA")

    array_faces = parafac_spatial_faces

    row = 0

    for col, data in enumerate(array_faces):
        worksheet2.write_column(row, col, data)

    ##

    worksheet3 = workbook.add_worksheet("Objects SSA")

    array_objects = parafac_spatial_objects

    row = 0

    for col, data in enumerate(array_objects):
        worksheet3.write_column(row, col, data)

    ##

    worksheet4 = workbook.add_worksheet("Factor Face SSA")

    array_f = factor_face

    row = 0

    for col, data in enumerate(array_f):
        worksheet4.write_column(row, col, data)

    ##

    worksheet5 = workbook.add_worksheet("Factor Object SSA")

    array_o = factor_obj

    row = 0

    for col, data in enumerate(array_o):
        worksheet5.write_column(row, col, data)

    ##

    print_peak_measures(faces_WSSA, title='Faces')

    f_l_right_P1, f_l_left_P1, f_l_right_N170, f_l_left_N170, f_a_right_P1, f_a_left_P1, f_a_right_N170, f_a_left_N170 = \
        data_per_hemisphere_6(faces_WSSA, title='Faces')

    f_l_right_P1_4, f_l_left_P1_4, f_l_right_N170_4, f_l_left_N170_4, f_a_right_P1_4, f_a_left_P1_4, f_a_right_N170_4, \
    f_a_left_N170_4 = data_per_hemisphere_4(faces_WSSA, title='Faces')

    print_peak_measures(objects_WSSA, title='Objects')

    o_l_right_P1, o_l_left_P1, o_l_right_N170, o_l_left_N170, o_a_right_P1, o_a_left_P1, o_a_right_N170, o_a_left_N170 = \
        data_per_hemisphere_6(objects_WSSA, title='Objects')

    o_l_right_P1_4, o_l_left_P1_4, o_l_right_N170_4, o_l_left_N170_4, o_a_right_P1_4, o_a_left_P1_4, o_a_right_N170_4, \
    o_a_left_N170_4 = data_per_hemisphere_4(objects_WSSA, title='Objects')

    worksheet6 = workbook.add_worksheet("ERPs without SSA")

    data = (['Name', name], ['Laterality', laterality], ['Group', group], ['Age', age], ['State', state], ['', ''],
            ['O1 Faces:', ''],
            ['P1 amplitude (µV)', faces_WSSA[4]["P1 amplitude"]], ['P1 latency (ms)', faces_WSSA[4]["P1 latency"]],
            ['N170 amplitude (µV)', faces_WSSA[4]["N170 amplitude"]],
            ['N170 latency (ms)', faces_WSSA[4]["N170 latency"]],
            ['', ''],
            ['O2 Faces:', ''],
            ['P1 amplitude (µV)', faces_WSSA[5]["P1 amplitude"]], ['P1 latency (ms)', faces_WSSA[5]["P1 latency"]],
            ['N170 amplitude (µV)', faces_WSSA[5]["N170 amplitude"]],
            ['N170 latency (ms)', faces_WSSA[5]["N170 latency"]],
            ['', ''],
            ['T5 Faces:', ''],
            ['P1 amplitude (µV)', faces_WSSA[2]["P1 amplitude"]], ['P1 latency (ms)', faces_WSSA[2]["P1 latency"]],
            ['N170 amplitude (µV)', faces_WSSA[2]["N170 amplitude"]],
            ['N170 latency (ms)', faces_WSSA[2]["N170 latency"]],
            ['', ''],
            ['T6 Faces:', ''],
            ['P1 amplitude (µV)', faces_WSSA[3]["P1 amplitude"]], ['P1 latency (ms)', faces_WSSA[3]["P1 latency"]],
            ['N170 amplitude (µV)', faces_WSSA[3]["N170 amplitude"]],
            ['N170 latency (ms)', faces_WSSA[3]["N170 latency"]],
            ['', ''],
            ['P3 Faces:', ''],
            ['P1 amplitude (µV)', faces_WSSA[0]["P1 amplitude"]], ['P1 latency (ms)', faces_WSSA[0]["P1 latency"]],
            ['N170 amplitude (µV)', faces_WSSA[0]["N170 amplitude"]],
            ['N170 latency (ms)', faces_WSSA[0]["N170 latency"]],
            ['', ''],
            ['P4 Faces:', ''],
            ['P1 amplitude (µV)', faces_WSSA[1]["P1 amplitude"]], ['P1 latency (ms)', faces_WSSA[1]["P1 latency"]],
            ['N170 amplitude (µV)', faces_WSSA[1]["N170 amplitude"]],
            ['N170 latency (ms)', faces_WSSA[1]["N170 latency"]],
            ['', ''],
            ['Faces Mean Latency per Hemisphere (6 electrodes):', ''],
            ['P1  left (ms)', f_l_left_P1], ['P1 right (ms)', f_l_right_P1],
            ['N170 left (ms)', f_l_left_N170], ['N170 right (ms)', f_l_right_N170],
            ['', ''],
            ['Faces Mean Amplitude per Hemisphere (6 electrodes):', ''],
            ['P1  left (µV)', f_a_left_P1], ['P1 right (µV)', f_a_right_P1],
            ['N170 left (µV)', f_a_left_N170], ['N170 right (µV)', f_a_right_N170],
            ['', ''], ['', ''],
            ['Faces Mean Latency per Hemisphere (4 electrodes):', ''],
            ['P1  left (ms)', f_l_left_P1_4], ['P1 right (ms)', f_l_right_P1_4],
            ['N170 left (ms)', f_l_left_N170_4], ['N170 right (ms)', f_l_right_N170_4],
            ['', ''],
            ['Faces Mean Amplitude per Hemisphere (4 electrodes):', ''],
            ['P1  left (µV)', f_a_left_P1_4], ['P1 right (µV)', f_a_right_P1_4],
            ['N170 left (µV)', f_a_left_N170_4], ['N170 right (µV)', f_a_right_N170_4],
            ['', ''], ['', ''],
            ['O1 Objects:', ''],
            ['P1 amplitude (µV)', objects_WSSA[4]["P1 amplitude"]], ['P1 latency (ms)', objects_WSSA[4]["P1 latency"]],
            ['N170 amplitude (µV)', objects_WSSA[4]["N170 amplitude"]],
            ['N170 latency (ms)', objects_WSSA[4]["N170 latency"]],
            ['', ''],
            ['O2 Objects:', ''],
            ['P1 amplitude (µV)', objects_WSSA[5]["P1 amplitude"]], ['P1 latency (ms)', objects_WSSA[5]["P1 latency"]],
            ['N170 amplitude (µV)', objects_WSSA[5]["N170 amplitude"]],
            ['N170 latency (ms)', objects_WSSA[5]["N170 latency"]],
            ['', ''],
            ['T5 Objects:', ''],
            ['P1 amplitude (µV)', objects_WSSA[2]["P1 amplitude"]], ['P1 latency (ms)', objects_WSSA[2]["P1 latency"]],
            ['N170 amplitude (µV)', objects_WSSA[2]["N170 amplitude"]],
            ['N170 latency (ms)', objects_WSSA[2]["N170 latency"]],
            ['', ''],
            ['T6 Objects:', ''],
            ['P1 amplitude (µV)', objects_WSSA[3]["P1 amplitude"]], ['P1 latency (ms)', objects_WSSA[3]["P1 latency"]],
            ['N170 amplitude (µV)', objects_WSSA[3]["N170 amplitude"]],
            ['N170 latency (ms)', objects_WSSA[3]["N170 latency"]],
            ['', ''],
            ['P3 Objects:', ''],
            ['P1 amplitude (µV)', objects_WSSA[0]["P1 amplitude"]], ['P1 latency (ms)', objects_WSSA[0]["P1 latency"]],
            ['N170 amplitude (µV)', objects_WSSA[0]["N170 amplitude"]],
            ['N170 latency (ms)', objects_WSSA[0]["N170 latency"]],
            ['', ''],
            ['P4 Objects:', ''],
            ['P1 amplitude (µV)', objects_WSSA[1]["P1 amplitude"]], ['P1 latency (ms)', objects_WSSA[1]["P1 latency"]],
            ['N170 amplitude (µV)', objects_WSSA[1]["N170 amplitude"]],
            ['N170 latency (ms)', objects_WSSA[1]["N170 latency"]],
            ['', ''],
            ['Objects Mean Latency per Hemisphere (6 electrodes):', ''],
            ['P1  left (ms)', o_l_left_P1], ['P1 right (ms)', o_l_right_P1],
            ['N170 left (ms)', o_l_left_N170], ['N170 right (ms)', o_l_right_N170],
            ['', ''],
            ['Objects Mean Amplitude per Hemisphere (6 electrodes):', ''],
            ['P1  left (µV)', o_a_left_P1], ['P1 right (µV)', o_a_right_P1],
            ['N170 left (µV)', o_a_left_N170], ['N170 right (µV)', o_a_right_N170],
            ['Objects Mean Latency per Hemisphere (4 electrodes):', ''],
            ['P1  left (ms)', o_l_left_P1_4], ['P1 right (ms)', o_l_right_P1_4],
            ['N170 left (ms)', o_l_left_N170_4], ['N170 right (ms)', o_l_right_N170_4],
            ['', ''],
            ['Objects Mean Amplitude per Hemisphere (4 electrodes):', ''],
            ['P1  left (µV)', o_a_left_P1_4], ['P1 right (µV)', o_a_right_P1_4],
            ['N170 left (µV)', o_a_left_N170_4], ['N170 right (µV)', o_a_right_N170_4])

    # Start from the first cell. Rows and columns are zero indexed.
    row = 0
    col = 0

    # Iterate over the data and write it out row by row.
    for name, score in data:
        worksheet6.write(row, col, name)
        worksheet6.write(row, col + 1, score)
        row += 1

    ##

    worksheet7 = workbook.add_worksheet("Faces without SSA")

    array_faces = parafac_spatial_WSSA_faces

    row = 0

    for col, data in enumerate(array_faces):
        worksheet7.write_column(row, col, data)

    ##

    worksheet8 = workbook.add_worksheet("Objects without SSA")

    array_objects = parafac_spatial_WSSA_objects

    row = 0

    for col, data in enumerate(array_objects):
        worksheet8.write_column(row, col, data)

    ##

    worksheet9 = workbook.add_worksheet("Factor Face without SSA")

    array_f = factor_face_WSSA

    row = 0

    for col, data in enumerate(array_f):
        worksheet9.write_column(row, col, data)

    ##

    worksheet10 = workbook.add_worksheet("Factor Object without SSA")

    array_o = factor_obj_WSSA

    row = 0

    for col, data in enumerate(array_o):
        worksheet10.write_column(row, col, data)

    workbook.close()
