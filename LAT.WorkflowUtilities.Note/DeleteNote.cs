using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

namespace LAT.WorkflowUtilities.Note
{
	public class DeleteNote : CodeActivity
	{
		[RequiredArgument]
		[Input("Note To Delete")]
		[ReferenceTarget("annotation")]
		public InArgument<EntityReference> NoteToDelete { get; set; }

		[Output("Was Note Deleted")]
		public OutArgument<bool> WasNoteDeleted { get; set; }

		protected override void Execute(CodeActivityContext executionContext)
		{
			ITracingService tracer = executionContext.GetExtension<ITracingService>();
			IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
			IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
			IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

			try
			{
				EntityReference noteToDelete = NoteToDelete.Get(executionContext);

				service.Delete("annotation", noteToDelete.Id);

				WasNoteDeleted.Set(executionContext, true);
			}
			catch (Exception e)
			{
				throw new InvalidPluginExecutionException(e.Message);
			}
		}
	}
}
