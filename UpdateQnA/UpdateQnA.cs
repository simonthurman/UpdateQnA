using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Microsoft.Azure.CognitiveServices.Knowledge.QnAMaker;
using Microsoft.Azure.CognitiveServices.Knowledge.QnAMaker.Models;
using System.Collections.Generic;

namespace UpdateQnA
{
    public static class UpdateQnA
    {
        [FunctionName("UpdateQnA")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            string question = req.Query["question"];
            string answer = req.Query["answer"];

            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            dynamic data = JsonConvert.DeserializeObject(requestBody);
            question = question ?? data?.question;
            answer = answer ?? data?.answer;

            string responseMessage = string.IsNullOrEmpty(question)
                ? "This HTTP triggered function executed successfully. Pass a QnA pair."
                : $"Your question {question}, and the answer {answer}.";

            var authoringKey = "";
            var sourceFile = "";
            var authoringUrl = $"";
            var kbId = "";

            var client = new QnAMakerClient(new ApiKeyServiceClientCredentials(authoringKey))
            { 
                Endpoint = authoringUrl 
            };

            //Delete Source File
            List<string> toDelete = new List<string>();

            toDelete.Add(sourceFile);

            var updateDelete = await client.Knowledgebase.UpdateAsync(kbId, new UpdateKbOperationDTO
            {
                Add = null,
                Update = null,
                Delete = new UpdateKbOperationDTODelete(null, toDelete)
            });

            updateDelete = await MonitorOperation(client, updateDelete);

            //Update Knowledge Base
            var updateKB = await client.Knowledgebase.UpdateAsync(kbId, new UpdateKbOperationDTO
            {
                Add = new UpdateKbOperationDTOAdd
                {
                    QnaList = new List<QnADTO> {
                        new QnADTO{
                            Questions = new List<string>
                            {
                                question
                            },
                            Answer = answer,
                            Source = sourceFile,
                            Metadata = new List<MetadataDTO>
                            {
                                new MetadataDTO {Name = "Category", Value = sourceFile},
                            }
                        }
                     },
                }

            }) ;

            updateKB = await MonitorOperation(client, updateKB);

            //Publish Knowledge Base
            await client.Knowledgebase.PublishAsync(kbId);

            return new OkObjectResult(responseMessage);
        }
        private static async Task<Operation> MonitorOperation(IQnAMakerClient client, Operation operation)
        {
            // Loop while operation is success
            for (int i = 0;
                i < 20 && (operation.OperationState == OperationStateType.NotStarted || operation.OperationState == OperationStateType.Running);
                i++)
            {
                Console.WriteLine("Waiting for operation: {0} to complete.", operation.OperationId);
                await Task.Delay(5000);
                operation = await client.Operations.GetDetailsAsync(operation.OperationId);
            }

            if (operation.OperationState != OperationStateType.Succeeded)
            {
                throw new Exception($"Operation {operation.OperationId} failed to completed.");
            }
            return operation;
        }
    }

}
