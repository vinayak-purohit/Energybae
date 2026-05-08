# Project Walkthrough: How I Built The Automation App

So here is the document about how I went through the whole task and what I actually did. 

### Planning Phase
First I started with the planning phase. When I read the document that was shared with the task details, I tried to map out how we will be doing the whole thing. I planned all the things and planned the walkthrough of how the app should work. The main idea was simple: take an electricity bill, read the data from it, and put that data into the Excel file. 
I planned what tools we will be using. We needed an AI to read the image, Python to handle the logic, and a way to edit the Excel template without messing up the formulas. 

### Creating Phase and Our Approach
Then I moved on to the creating part. I started by writing the backend code. My approach was to first get the AI to successfully pull the data from the bill. The logic behind what we did was to send the image to the AI, ask it for specific fields like name and monthly units, and get a JSON response back. 
After that, I used Python to open the `Energybae_Customer_Proposal.xlsx` file. The logic here was to just insert the values into the exact cells where they belong and save it as a new file. This way the original template stays safe and all the calculation formulas keep working.

### Problems Faced and Fixes
Things were not entirely smooth, I faced a bunch of problems while making it.

1. **API Limits:** While using the Gemini API, there were some problems. It was showing out of limits errors when I was testing. I couldn't keep working with it stopping all the time. So I tried different APIs to see what works better. I finally settled on the Groq APIs. They were really fast and were giving the answers without stopping me.

2. **AI Reading Wrong Values:** Then I had another problem. The AI was reading the image properly for normal text, but there were issues where it couldn't read the amount properly. It was getting the consumption values wrong and giving wrong output. Initially I was using just a single image of the bill. Then I tried fixing it by switching the approach. I changed the code so if we use more images, like multiple pages, we can get more accurate answers. So I used more images in the prompt and it finally started getting the numbers right.

3. **Excel File Output:** I had to figure out how to use the Excel to give the output properly. I made sure the code maps the AI output to the exact cell numbers in the Excel template so the formulas would calculate the final solar load automatically. 

4. **Changing the UI:** Initially we were using Streamlit for the UI of the app because it was easy to set up. But it didn't look that great. So then we switched to a proper frontend. I built it with HTML, CSS, and JS and we got our inspiration from the original website of EnergyBae. I used their green and blue colors and style to make it look like a real professional app instead of just a script. Then we linked this frontend to the backend.

### Future Plans for Improvement

Right now, this is a solid working prototype. But if I get the chance to keep working on this, there are a few things I would improve.

1. **WhatsApp / Zapier Integration:** First, I would love to connect this directly to WhatsApp or Zapier. That way, a sales rep could literally just forward a PDF bill on WhatsApp and receive the filled Excel sheet back in seconds. No need to even open the website. It would make the whole workflow so much faster for the team on the ground.

2. **Database for Proposal History:** I would also add a database to save these proposals so you have a history of all leads. Right now, every time you generate a proposal it just downloads and that's it. But if we store the extracted data and the generated files in a database, the sales team can go back and check old proposals, track which customers they already contacted, and basically use it as a mini CRM for solar leads.

3. **Image Pre-processing for Better Accuracy:** And the things that I am planning for the future is to improve the output of the AI even more. I was thinking we can do one thing: while processing the images before giving it to the AI, we can create a Python script which will extract or divide the important parts of the bill. Like extracting only the part which is important and contains all the necessary details. So the flow would be: we get the file, we process it using the Python script to cut out only the part that is needed, and then we directly give that cropped part to the AI. This way the AI doesn't get confused by extra text and it will really improve the final output.

4. **PDF Proposal Generation:** Along with the Excel file, we could also auto-generate a professional PDF proposal with the EnergyBae branding that can be directly shared with the customer. That would make the whole thing look even more polished.

And yeah, that is basically how we did it, how we solved the problems that came up, and what we are planning next!
