'use strict';

(function () {

	// The initialize function is run each time the page is loaded.
	Office.initialize = function (reason) {
		$(document).ready(function () {

			if (Office.context.requirements.isSetSupported('WordApi', 1.1)) {
				$('#edit').click(BindEmpDetails);
				$('#update').click(SaveEmpDetails);
				$('#ddlEmployee').change(ClearDocumnet);
			}
			Word.run(function (context) {

				const Paragraph = context.document.body;
				Paragraph.font.set({
					name: "Calibri",
					bold: false,
					size: 11,
					color: "Black",
				});
				context.document.body.clear();

				return context.sync();
			})

		});
	};


	function BindEmpDetails() {
		Word.run(function (context) {

			
			var thisDocument = context.document;

			var range = thisDocument.getSelection();

			range.insertText('Edit Employee Details.\n', Word.InsertLocation.replace);

			const Paragraph = context.document.body.paragraphs.getFirst();
			Paragraph.font.set({
				name: "Courier New",
				bold: true,
				size: 18,
				color: "Blue",
			});
			//context.sync();

			var e = document.getElementById("ddlEmployee");
			var id = e.options[e.selectedIndex].value;
			//debugger;
			if (id == null || id == undefined) {
				return context.sync();
			}
			const secondParagraph = context.document.body.paragraphs.getFirst().getNext();
			var result = "hello";
			$.ajax({
				url: '../../api/Employee',
				type: 'GET',
				data: {
					empid: id
				},
				contentType: 'application/json;charset=utf-8'
			}).done(function (data) {
				result = data;
				const tableData = [];
				//debugger;
				//var aa = result[0].length;
				for (var i = 0; i < result.length; i++) {
					tableData.push([result[i].employee_Name, result[i].email]);
				}
				secondParagraph.insertTable(data.length, 2, "", tableData);
				secondParagraph.font.set({
					name: "Courier New",
					bold: true,
					size: 12,
					color: "Red",
				});
				
				e.options[e.selectedIndex].value = id;
				return context.sync();
			}).fail(function (status) {
				result = "Could not communicate with the server.";
			});

			return context.sync();
		})
			.catch(function (error) {
				console.log('Error: ' + JSON.stringify(error));
				alert('Error: ' + JSON.stringify(error));
				if (error instanceof OfficeExtension.Error) {
					console.log('Debug info: ' + JSON.stringify(error.debugInfo));
				}
			});
	}


	function SaveEmpDetails() {
		Word.run(function (context) {


			var e = document.getElementById("ddlEmployee");
			var id = e.options[e.selectedIndex].value;


			if (id == null || id == undefined) {
				return context.sync();
			}
			var EmpData = context.document.body;
			context.load(EmpData, 'text');


			return context.sync()
				.then(function () {
					// Get the longest word from the selection.
					var words = EmpData.text.split("Employee Name");

					var wordsN = words[1].split("Joining Date");
					var Ename = wordsN[0].toString().trim('\t').trim('\r\n');

					var wordsJ = wordsN[1].split("Department Name");
					var JoingDate = wordsJ[0].toString().trim('\t').trim('\r\n');

					var wordsD = wordsJ[1].split("Email");
					var Department = wordsD[0].toString().trim('\t').trim('\r\n');

					var wordsE = wordsD[1].split("Address");
					var Email = wordsE[0].toString().trim('\t').trim('\r\n');

					var wordsA = wordsE[1].split("Mobile No");
					var Address = wordsA[0].toString().trim('\t').trim('\r\n');
					var Mobile = wordsA[1].toString().trim('\t').trim('\r\n');

				
					var result = "";
					$.ajax({
						url: '../../api/Employee/empid',
						type: 'GET',
						data: {
							empid: id, name: Ename, date: JoingDate, depart: Department, emails: Email, add: Address, mobileno: Mobile
						},
						contentType: 'application/json;charset=utf-8'
					}).done(function (data) {
						result = data;

						var range = context.document.body.paragraphs.getLast();
						if (result == 1) {
							range.insertText('Employee Update Successfully.\n', Word.InsertLocation.replace);
						}
						else {
							range.insertText('Employee Update Not Successfully.\n', Word.InsertLocation.replace);
						}
						const Paragraph = context.document.body.paragraphs.getLast();
						Paragraph.font.set({
							name: "Courier New",
							bold: true,
							size: 12,
							color: "Black",
						});

					//	
						return context.sync();
					}).fail(function (status) {
						result = "Could not communicate with the server.";
					});

				})
				.then(context.sync);


			
		})
			.catch(function (error) {
				console.log('Error: ' + JSON.stringify(error));
				alert('Error: ' + JSON.stringify(error));
				if (error instanceof OfficeExtension.Error) {
					console.log('Debug info: ' + JSON.stringify(error.debugInfo));
				}
			});
	}

	function ClearDocumnet() {
		Word.run(function (context) {


			const Paragraph = context.document.body;
			Paragraph.font.set({
				name: "Calibri",
				bold: false,
				size: 11,
				color: "Black",
			});
			context.document.body.clear();
			
			return context.sync();
		})
			.catch(function (error) {
				console.log('Error: ' + JSON.stringify(error));
				alert('Error: ' + JSON.stringify(error));
				if (error instanceof OfficeExtension.Error) {
					console.log('Debug info: ' + JSON.stringify(error.debugInfo));
				}
			});
	}


})();