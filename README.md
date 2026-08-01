# BdResNet Dataset

A dataset collection script that fetches research paper metadata from [OpenAlex](https://openalex.org/) for Bangladeshi universities and institutions. Each institution's data is saved as a separate sheet in a single Excel workbook, and automatically synced with a shared Google Drive folder.

## 📦 Requirements

- Python 3.8+
- Internet connection (for OpenAlex API)
- Google Cloud account (for Drive sync)

## 🚀 Setup & Usage

### 1. Clone the Repository

```bash
git clone <repository-url>
cd BdResNet-Dataset
```

### 2. Install Dependencies

```bash
pip install -r requirements.txt
```

### 3. Create the `.env` File

**Windows:**
```bash
copy .env.example .env
```

**Mac / Linux:**
```bash
cp .env.example .env
```

### 4. Fill in `.env`

```env
OPENALEX_EMAIL=yourname@gmail.com
GOOGLE_SERVICE_ACCOUNT_JSON=path/to/service_account.json
GOOGLE_DRIVE_FOLDER_ID=your_folder_id_here
```

> **Why is email required?**  
> OpenAlex places users who provide an email in the "Polite Pool", which grants a significantly higher rate limit. This is completely free.

### 5. Run the Script

```bash
python script.py
```

**What happens each run:**
1. 📥 Downloads the latest Excel from Google Drive
2. ⏭️ Skips institutions already present in the file
3. 📡 Fetches data only for new institutions
4. 📤 Uploads the updated Excel back to Drive

---

## ⚙️ Configuration

You can add or remove institutions in `script.py`:

```python
INSTITUTIONS = {
    "Green University of Bangladesh": "I2802576217",
    "American International University - Bangladesh": "I103434671",
    "North South University": "I157386601",
    # Add more institutions here
}
```

To find an institution's OpenAlex ID, use the [OpenAlex Institution Search](https://openalex.org/institutions).

You can also control how many papers are fetched per institution:

```python
process_institutions(
    ...
    max_papers_per_inst=100   # Set to None to fetch all papers
)
```

---

## 🔒 Security

- Never commit your `.env` file to Git (it is already listed in `.gitignore`)
- Never commit `service_account.json` to Git (`*.json` is in `.gitignore`)
- Only commit `.env.example` as a template — never the actual `.env`
- Share credentials with team members through a secure channel only

---

## 📄 License

This project is for academic research purposes.

## All Universities of Bangladesh

1. University of Dhaka
2. University of Rajshahi
3. University of Chittagong
4. Jahangirnagar University
5. Islamic University, Bangladesh
6. Khulna University
7. Jagannath University
8. Comilla University
9. Jatiya Kabi Kazi Nazrul Islam University
10. Bangladesh University of Professionals
11. Begum Rokeya University
12. University of Barisal
13. Rabindra University, Bangladesh
14. Netrokona University
15. Kishoreganj University
16. Meherpur University
17. University of Bogura
18. Thakurgaon University
19. Shahjalal University of Science and Technology
20. Hajee Mohammad Danesh Science & Technology University
21. Mawlana Bhashani Science and Technology University
22. Patuakhali Science and Technology University
23. Noakhali Science and Technology University
24. Jashore University of Science and Technology
25. Pabna University of Science and Technology
26. Gopalganj Science and Technology University
27. Rangamati Science and Technology University
28. Jamalpur Science and Technology University
29. Chandpur Science and Technology University
30. Sunamganj Science and Technology University
31. Lakshmipur Science and Technology University
32. Pirojpur Science and Technology University
33. Satkhira University of Science and Technology
34. Narayanganj Science and Technology University
35. Bangladesh University of Engineering & Technology
36. Military Institute of Science and Technology
37. Rajshahi University of Engineering & Technology
38. Khulna University of Engineering & Technology
39. Chittagong University of Engineering & Technology
40. Dhaka University of Engineering & Technology
41. Bangladesh Agricultural University
42. Gazipur Agricultural University
43. Sher-e-Bangla Agricultural University
44. Sylhet Agricultural University
45. Khulna Agricultural University
46. Habiganj Agricultural University
47. Kurigram Agricultural University
48. Shariatpur Agriculture University
49. Bangladesh Medical University
50. Chittagong Medical University
51. Rajshahi Medical University
52. Sylhet Medical University
53. Khulna Medical University
54. Chittagong Veterinary and Animal Sciences University
55. Bangladesh University of Textiles
56. Bangladesh Maritime University
57. University of Frontier Technology, Bangladesh
58. Aviation and Aerospace University Bangladesh
59. National University Bangladesh
60. Bangladesh Open University
61. Islamic Arabic University
62. Dhaka Central University
63. International University of Business Agriculture and Technology
64. North South University
65. Independent University, Bangladesh
66. American International University-Bangladesh
67. Dhaka International University
68. International Islamic University, Chittagong
69. Asian University of Bangladesh
70. East West University
71. Gono Bishwabidyalay
72. People's University of Bangladesh
73. Queens University
74. University of Asia Pacific (Bangladesh)
75. Chittagong Independent University
76. Bangladesh University
77. BGC Trust University Bangladesh
78. Brac University
79. Manarat International University
80. Premier University, Chittagong
81. Southern University, Bangladesh
82. Sylhet International University
83. University of Development Alternative
84. City University, Bangladesh
85. Daffodil International University
86. Green University of Bangladesh
87. IBAIS University
88. Leading University
89. Northern University, Bangladesh
90. Prime University
91. Southeast University
92. Stamford University Bangladesh
93. State University of Bangladesh
94. Eastern University, Bangladesh
95. Metropolitan University
96. Millennium University
97. Primeasia University
98. Royal University of Dhaka
99. United International University
100. University of Information Technology and Sciences
101. University of South Asia, Bangladesh
102. Presidency University
103. Uttara University
104. Victoria University of Bangladesh
105. World University of Bangladesh
106. Asa University Bangladesh
107. Bangladesh Islami University
108. East Delta University
109. Northern University of Business and Technology Khulna
110. Britannia University
111. Feni University
112. Khwaja Yunus Ali University
113. European University of Bangladesh
114. First Capital University Of Bangladesh
115. BGMEA University of Fashion & Technology
116. Hamdard University Bangladesh
117. Ishakha International University
118. North East University Bangladesh
119. North Western University, Bangladesh
120. Port City International University
121. Varendra University
122. Sonargaon University
123. Cox's Bazar International University
124. Fareast International University
125. German University Bangladesh
126. North Bengal International University
127. Notre Dame University Bangladesh
128. Ranada Prasad Shaha University
129. Brahmaputra International University
130. Times University Bangladesh
131. Canadian University of Bangladesh
132. Global University Bangladesh
133. NPI University of Bangladesh
134. Rabindra Maitree University
135. University of Scholars
136. University of Creative Technology Chittagong
137. Anwer Khan Modern University
138. University of Global Village
139. Khulna Khan Bahadur Ahsanullah University
140. Trust University, Barishal
141. University of Brahmanbaria
142. University of Skill Enrichment and Technology
143. International Standard University
144. ZNRF University of Management Sciences
145. Bandarban University
146. RTM Al-Kabir Technical University
147. International Islami University of Science and Technology Bangladesh
148. Microland University of Science and Technology
149. Lalon University of Science & Arts
150. Grameen University
151. University of Science & Technology Chittagong
152. Ahsanullah University of Science and Technology
153. Pundra University of Science and Technology
154. Bangladesh University of Business and Technology
155. Atish Dipankar University of Science and Technology
156. ZH Sikder University of Science & Technology
157. Rajshahi Science & Technology University
158. Bangladesh Army International University of Science & Technology
159. Bangladesh Army University of Science & Technology, Khulna
160. Bangladesh Army University of Science and Technology, Saidpur
161. CCN University of Science & Technology
162. Central University of Science and Technology
163. Dr. Momtaz Begum University of Science and Technology
164. Central Women's University
165. Shanto-Mariam University of Creative Technology
166. University of Liberal Arts Bangladesh
167. Bangladesh University of Health Sciences
168. Exim Bank Agricultural University Bangladesh
169. Bangladesh Army University of Engineering & Technology
170. Tagore University of Creative Arts
171. Islamic University of Technology
172. Asian University for Women
