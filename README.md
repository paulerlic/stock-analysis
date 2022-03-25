# Green Energy Stocks

### Overview 

We are helping steve and his parents determine the performance of "green energy stocks". Our goal is to automate the analysis process so Steve can regularly analyze this group of stocks. 

### Purpose

The goal of this project was to perform an annual return analysis on green energy stocks by using VBA to automate the analysis process, we will start with the provided starter code and attempt to refactor it in order to speed up the run times of the script.    

### Results

Our stocks did better in 2017 overall relative to 2018 (11 of 12) stocks showing positive returns for the year. Whereas in 2018 only two stocks had positive returns. 

  * The original script posted a run time of (0.8750 seconds 2017) and (0.8516 seconds 2018) to run.

  * Refactoring led to a significant time savings, now the script only takess (0.1875 seconds 2017) and (0.1797 seconds 2018) to run, This is nearly 5x faster than the original code.


![OG:2017](https://user-images.githubusercontent.com/86337035/160211584-9a2e0d3e-a117-4d8b-a1a8-d1b87af80fa0.png)


![OG:2018](https://user-images.githubusercontent.com/86337035/160211632-7e0fa8b3-38fa-49f9-900b-c0820e959515.png)

![Refac:2017](https://user-images.githubusercontent.com/86337035/160211658-8f7c72ea-60ce-4579-a4a9-f8ffcec6d028.png)

![Refac:2018](https://user-images.githubusercontent.com/86337035/160211696-48132adc-d1c3-4d36-bfe5-5763a36be521.png)

### Summary

* What are the advantages or disadvantages of refactoring code?

  I. Advantages of refactoring : First and foremost, speed refactoring the code sped up the run time of our script almost five fold. Granted we were working with a relativley small data set. If we were dealing with a larger data set this would be extremely important and could mean the difference between this script executing in 1 hour or 5 hours. Other upsides are organziation, ease of updating, debugging. With properly organized code mundane tasks like updating and sharing code become easier because not only do you understand what the code is doing, when you share other people will as well if you did a good job with the organization of the code. 
 
  II. Disadvantages of refactoring : Refactoring the code can lead to new potential bugs. This means more time working on one particular task and in a business enviroment where time is money sometimes it is best to follow the age old adage of " if it ain't broke don't fix it".  

* How do these pros and cons apply to refactoring the original VBA script?

  1. Pros:  
      * Indexing /output arrays allow for easier updating 
      * More efficiency F
      * Faster run times.

  2. Cons:
      * It took me several hours to refactor the original code and we only saved 0.69 seconds in run time.
