using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

namespace LAT.WorkflowUtilities.Note
{
	public class CopyNote : CodeActivity
	{
		[RequiredArgument]
		[Input("Note To Copy")]
		[ReferenceTarget("annotation")]
		public InArgument<EntityReference> NoteToCopy { get; set; }

		[Input("Record Dynamic Url")]
		[RequiredArgument]
		public InArgument<string> RecordUrl { get; set; }

		[RequiredArgument]
		[Default("True")]
		[Input("Copy Attachment?")]
		public InArgument<bool> CopyAttachment { get; set; }

		[Output("Was Note Copied")]
		public OutArgument<bool> WasNoteCopied { get; set; }
		protected override void Execute(CodeActivityContext executionContext)
		{
			ITracingService tracer = executionContext.GetExtension<ITracingService>();
			IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
			IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
			IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

			try
			{
				EntityReference noteToCopy = NoteToCopy.Get(executionContext);
				string recordUrl = RecordUrl.Get<string>(executionContext);
				bool copyAttachment = CopyAttachment.Get(executionContext);

				var dup = new DynamicUrlParser(recordUrl);

				string newEntityLogical = dup.GetEntityLogicalName(service);

				Entity note = GetNote(service, noteToCopy.Id);
				if (note.GetAttributeValue<EntityReference>("objectid").Id == dup.Id && note.GetAttributeValue<EntityReference>("objectid").LogicalName == newEntityLogical)
				{
					WasNoteCopied.Set(executionContext, false);
					return;
				}

				Entity newNote = new Entity("annotation");
				newNote["objectid"] = new EntityReference(newEntityLogical, dup.Id);
				newNote["notetext"] = note.GetAttributeValue<string>("notetext");
				newNote["subject"] = note.GetAttributeValue<string>("subject");
				if (copyAttachment)
				{
					newNote["isdocument"] = note.GetAttributeValue<bool>("isdocument");
					newNote["filename"] = note.GetAttributeValue<string>("filename");
					newNote["filesize"] = note.GetAttributeValue<int>("filesize");
					newNote["documentbody"] = note.GetAttributeValue<string>("documentbody");
				}
				else
					newNote["isdocument"] = false;

				service.Create(newNote);

				WasNoteCopied.Set(executionContext, true);
			}
			catch (Exception e)
			{
				throw new InvalidPluginExecutionException(e.Message);
			}
		}

		private Entity GetNote(IOrganizationService service, Guid noteId)
		{
			return service.Retrieve("annotation", noteId, new ColumnSet("objectid", "documentbody", "filename", "filesize", "isdocument", "notetext", "subject"));
		}
	}
}
