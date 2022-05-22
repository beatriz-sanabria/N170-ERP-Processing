# N170-ERP-Processing

This project aimed is to apply the tensor decomposition, CANDECOMP/PARAFAC (CP), to extract Event Related Potentials features, and the singular spectrum analysis (SSA) to remove the undesired brain activities and artifacts and to find the amplitude and latency measures of P1 and N170.

Before applying CP decomposition, EEG needed to be transformed into a tensor. In the first step, we grouped the round and oval faces and did the same for the objects. Then, we created epochs per stimulus considering 100 ms previous and 500 ms after it, obtaining 128 epochs in total. With this information, a three-dimensional tensor was created for faces and objects and was made up of the dimensions: channel, time, and trial. With that, the CP decomposition can explore spatial, temporal, and spectral (across trials) correlations in the data to find the ERPs features.

After the CP decomposition, we used the SSA method to remove the beta activity due to attention and remaining artifacts from the time components. After that, the time and spatial component arrays and choose the selected component in the previous step to visualize the ERPs in the channels of interest (T5, T6, P3, P4, O1, and O2), and the amplitude and latency of P1 and N170 were calculated automatically.

The code expects as input the electroencephalogram recorded during the patient saw an emotions task in .cnt format. As output, the graphs related to the ERPs, P1, and N170, are shown, and an excel file is generated with the patient information and all the data printed in the console. Furthermore, during the process, you can see the results obtained after applying CP decomposition and the two types of filtering proposed. The following images are examples of the outputs.



![N170_1](https://user-images.githubusercontent.com/60671532/169672925-d9e9730e-b76d-45b7-8ddb-e5df4cf7eb40.png)
![N170_2](https://user-images.githubusercontent.com/60671532/169672926-aceb3e07-2121-46c1-baae-e78b10956b2b.png)
![N170_3](https://user-images.githubusercontent.com/60671532/169672927-b78cfc7f-3c80-4d1c-88b9-6561ca348528.png)


*** The purpose of this repository is to show the main file and the results we obtained after running it. 
    All the codes were programmed in Python.
    
    Due to the code license, we did not make a continuous versioning 
    in the repository
