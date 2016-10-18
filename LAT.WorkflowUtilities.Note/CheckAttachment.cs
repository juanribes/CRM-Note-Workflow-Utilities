using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

namespace LAT.WorkflowUtilities.Note
{
	public class CheckAttachment : CodeActivity
	{
		[RequiredArgument]
		[Input("Note To Check")]
		[ReferenceTarget("annotation")]
		public InArgument<EntityReference> NoteToCheck { get; set; }

		[Output("Has Attachment")]
		public OutArgument<bool> HasAttachment { get; set; }

		protected override void Execute(CodeActivityContext executionContext)
		{
			ITracingService tracer = executionContext.GetExtension<ITracingService>();
			IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
			IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
			IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

			try
			{
				EntityReference noteToCheck = NoteToCheck.Get(executionContext);

				HasAttachment.Set(executionContext, CheckForAttachment(service, noteToCheck.Id));
			}
			catch (Exception ex)
			{
				tracer.Trace("Exception: {0}", ex.ToString());
			}
		}

		private static bool CheckForAttachment(IOrganizationService service, Guid noteId)
		{
			Entity note = service.Retrieve("annotation", noteId, new ColumnSet("isdocument"));

			object oIsDocument;
			bool hasValue = note.Attributes.TryGetValue("isdocument", out oIsDocument);
			if (!hasValue)
				return false;

			return (bool)oIsDocument;
		}
	}
}
