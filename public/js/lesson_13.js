const questions = [
	{
		id: 5.1_1,
		question: "",
		checkFunction: function () {
			//document.write(`<div>${}</div><div><a href="../components/lesson9.html">Повернутися до уроку</a></div>`)
		}
	},
	// ========================================================================================================================================
	{
		id: 5.2_1,
		question: "",
		checkFunction: function () {
			//document.write(`<div>${}</div><div><a href="../components/lesson9.html">Повернутися до уроку</a></div>`)
		}
	},
	// ========================================================================================================================================
	{
		id: 5.3_2,
		question: "",
		checkFunction: function () {
			//document.write(`<div>${}</div><div><a href="../components/lesson9.html">Повернутися до уроку</a></div>`)
		}
	},
];

function showQuestion(questionId) {
	const questionObj = questions.find(question => question.id === questionId);
	if (questionObj) {
		questionObj.checkFunction();
	} else {
		document.getElementById('question-content').innerText = "Це питання відсутнє";
	}
}