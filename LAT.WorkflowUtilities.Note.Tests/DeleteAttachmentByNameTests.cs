using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using Moq;
using System;
using System.Activities;
using System.Collections.Generic;

namespace LAT.WorkflowUtilities.Note.Tests
{
	[TestClass]
	public class DeleteAttachmentByNameTests
	{
		#region Class Constructor
		private readonly string _namespaceClassAssembly;
		public DeleteAttachmentByNameTests()
		{
			//[Namespace.class name, assembly name] for the class/assembly being tested
			//Namespace and class name can be found on the class file being tested
			//Assembly name can be found under the project properties on the Application tab
			_namespaceClassAssembly = "LAT.WorkflowUtilities.Note.DeleteAttachmentByName" + ", " + "LAT.WorkflowUtilities.Note";
		}
		#endregion
		#region Test Initialization and Cleanup
		// Use ClassInitialize to run code before running the first test in the class
		[ClassInitialize()]
		public static void ClassInitialize(TestContext testContext) { }

		// Use ClassCleanup to run code after all tests in a class have run
		[ClassCleanup()]
		public static void ClassCleanup() { }

		// Use TestInitialize to run code before running each test 
		[TestInitialize()]
		public void TestMethodInitialize() { }

		// Use TestCleanup to run code after each test has run
		[TestCleanup()]
		public void TestMethodCleanup() { }
		#endregion

		[TestMethod]
		public void NoMatch()
		{
			//Target
			Entity targetEntity = null;

			//Input parameters
			var inputs = new Dictionary<string, object>
			{
				{ "NoteWithAttachment", new EntityReference("annotation", Guid.NewGuid()) },
				{ "FileName", "test.txt"},
				{ "AppendNotice", false }
			};

			//Expected value
			const int expected = 0;

			//Invoke the workflow
			var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, NoMatchSetup);

			//Test
			Assert.AreEqual(expected, output["NumberOfAttachmentsDeleted"]);
		}

		/// <summary>
		/// Modify to mock CRM Organization Service actions
		/// </summary>
		/// <param name="serviceMock">The Organization Service to mock</param>
		/// <returns>Configured Organization Service</returns>
		private static Mock<IOrganizationService> NoMatchSetup(Mock<IOrganizationService> serviceMock)
		{
			Entity note1 = new Entity("annotation");
			note1["filesize"] = 5000;
			note1["filename"] = "text.docx";
			note1["isdocument"] = true;

			serviceMock.Setup(t =>
				t.Retrieve(It.IsAny<string>(), It.IsAny<Guid>(), It.IsAny<ColumnSet>()))
				.ReturnsInOrder(note1);

			return serviceMock;
		}

		[TestMethod]
		public void OneMatch()
		{
			//Target
			Entity targetEntity = null;

			//Input parameters
			var inputs = new Dictionary<string, object>
			{
				{ "NoteWithAttachment", new EntityReference("annotation", Guid.NewGuid()) },
				{ "FileName", "test.txt"},
				{ "AppendNotice", false }
			};

			//Expected value
			const int expected = 1;

			//Invoke the workflow
			var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, OneMatchSetup);

			//Test
			Assert.AreEqual(expected, output["NumberOfAttachmentsDeleted"]);
		}

		/// <summary>
		/// Modify to mock CRM Organization Service actions
		/// </summary>
		/// <param name="serviceMock">The Organization Service to mock</param>
		/// <returns>Configured Organization Service</returns>
		private static Mock<IOrganizationService> OneMatchSetup(Mock<IOrganizationService> serviceMock)
		{
			Entity note1 = new Entity("annotation");
			note1["filesize"] = 5000;
			note1["filename"] = "test.txt";
			note1["isdocument"] = true;

			serviceMock.Setup(t =>
				t.Retrieve(It.IsAny<string>(), It.IsAny<Guid>(), It.IsAny<ColumnSet>()))
				.ReturnsInOrder(note1);

			return serviceMock;
		}

		[TestMethod]
		public void DeleteZeroNotDocument()
		{
			//Target
			Entity targetEntity = null;

			//Input parameters
			var inputs = new Dictionary<string, object>
			{
				{ "NoteWithAttachment", new EntityReference("annotation", Guid.NewGuid()) },
				{ "FileName", "test.txt"},
				{ "AppendNotice", false }
			};

			//Expected value
			const int expected = 0;

			//Invoke the workflow
			var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, DeleteZeroNotDocumentSetup);

			//Test
			Assert.AreEqual(expected, output["NumberOfAttachmentsDeleted"]);
		}

		/// <summary>
		/// Modify to mock CRM Organization Service actions
		/// </summary>
		/// <param name="serviceMock">The Organization Service to mock</param>
		/// <returns>Configured Organization Service</returns>
		private static Mock<IOrganizationService> DeleteZeroNotDocumentSetup(Mock<IOrganizationService> serviceMock)
		{
			Entity note1 = new Entity("annotation");
			note1["isdocument"] = false;

			serviceMock.Setup(t =>
				t.Retrieve(It.IsAny<string>(), It.IsAny<Guid>(), It.IsAny<ColumnSet>()))
				.ReturnsInOrder(note1);

			return serviceMock;
		}

		/// <summary>
		/// Invokes the workflow.
		/// </summary>
		/// <param name="name">Namespace.Class, Assembly</param>
		/// <param name="target">The target entity</param>
		/// <param name="inputs">The workflow input parameters</param>
		/// <param name="configuredServiceMock">The function to configure the Organization Service</param>
		/// <returns>The workflow output parameters</returns>
		private static IDictionary<string, object> InvokeWorkflow(string name, ref Entity target, Dictionary<string, object> inputs,
			Func<Mock<IOrganizationService>, Mock<IOrganizationService>> configuredServiceMock)
		{
			var testClass = Activator.CreateInstance(Type.GetType(name)) as CodeActivity;

			var serviceMock = new Mock<IOrganizationService>();
			var factoryMock = new Mock<IOrganizationServiceFactory>();
			var tracingServiceMock = new Mock<ITracingService>();
			var workflowContextMock = new Mock<IWorkflowContext>();

			//Apply configured Organization Service Mock
			if (configuredServiceMock != null)
				serviceMock = configuredServiceMock(serviceMock);

			IOrganizationService service = serviceMock.Object;

			//Mock workflow Context
			var workflowUserId = Guid.NewGuid();
			var workflowCorrelationId = Guid.NewGuid();
			var workflowInitiatingUserId = Guid.NewGuid();

			//Workflow Context Mock
			workflowContextMock.Setup(t => t.InitiatingUserId).Returns(workflowInitiatingUserId);
			workflowContextMock.Setup(t => t.CorrelationId).Returns(workflowCorrelationId);
			workflowContextMock.Setup(t => t.UserId).Returns(workflowUserId);
			var workflowContext = workflowContextMock.Object;

			//Organization Service Factory Mock
			factoryMock.Setup(t => t.CreateOrganizationService(It.IsAny<Guid>())).Returns(service);
			var factory = factoryMock.Object;

			//Tracing Service - Content written appears in output
			tracingServiceMock.Setup(t => t.Trace(It.IsAny<string>(), It.IsAny<object[]>())).Callback<string, object[]>(MoqExtensions.WriteTrace);
			var tracingService = tracingServiceMock.Object;

			//Parameter Collection
			ParameterCollection inputParameters = new ParameterCollection { { "Target", target } };
			workflowContextMock.Setup(t => t.InputParameters).Returns(inputParameters);

			//Workflow Invoker
			var invoker = new WorkflowInvoker(testClass);
			invoker.Extensions.Add(() => tracingService);
			invoker.Extensions.Add(() => workflowContext);
			invoker.Extensions.Add(() => factory);

			return invoker.Invoke(inputs);
		}
	}
}
