using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

namespace LAT.WorkflowUtilities.Note
{
	public class UpdateNoteText : CodeActivity
	{
		[RequiredArgument]
		[Input("Note To Update")]
		[ReferenceTarget("annotation")]
		public InArgument<EntityReference> NoteToUpdate { get; set; }

		[RequiredArgument]
		[Input("New Text")]
		public InArgument<string> NewText { get; set; }

		[Output("Updated Text")]
		public OutArgument<string> UpdatedText{ get; set; }
		protected override void Execute(CodeActivityContext executionContext)
		{
			ITracingService tracer = executionContext.GetExtension<ITracingService>();
			IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
			IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
			IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

			try
			{
				EntityReference noteToUpdate = NoteToUpdate.Get(executionContext);
				string newText = NewText.Get(executionContext);

				Entity note = new Entity("annotation");
				note.Id = noteToUpdate.Id;
				note["notetext"] = newText;
				service.Update(note);

				UpdatedText.Set(executionContext, newText);
			}
			catch (Exception e)
			{
				throw new InvalidPluginExecutionException(e.Message);
			}
		}
	}
}
