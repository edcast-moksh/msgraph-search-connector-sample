using Microsoft.Graph;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using Newtonsoft.Json.Serialization;
using System;
using System.IO;

namespace PartsInventoryConnector.MicrosoftGraph
{
    // The Microsoft Graph SDK serializes enumerations in camelCase.
    // The Microsoft Graph service currently requires the PropertyType enum
    // to be PascalCase. This will override the Microsoft Graph serialization
    // If the Microsoft Graph service changes to accept camelCase this will no
    // longer be necessary.
    class CustomContractResolver : DefaultContractResolver
    {
        protected override JsonConverter ResolveContractConverter(Type objectType)
        {
            if (typeof(PropertyType).IsAssignableFrom(objectType))
            {
                // This default converter uses PascalCase
                return new StringEnumConverter();
            }
            return base.ResolveContractConverter(objectType);
        }
    }

    // In order to hook up the custom contract resolver for
    // PropertyType, we need to implement a custom serializer to
    // pass to the MicrosoftGraphServiceClient.
    public class CustomSerializer : ISerializer
    {

        private Serializer _microsoftGraphSerializer;
        private JsonSerializerSettings _jsonSerializerSettings;

        public CustomSerializer()
        {
            _microsoftGraphSerializer = new Serializer();

            _jsonSerializerSettings = new JsonSerializerSettings
            {
                ContractResolver = new CustomContractResolver()
            };
        }

        // For deserialize, just pass through to the default
        // Microsoft Graph SDK serializer
        public T DeserializeObject<T>(Stream stream)
        {
            return _microsoftGraphSerializer.DeserializeObject<T>(stream);
        }

        // For deserialize, just pass through to the default
        // Microsoft Graph SDK serializer
        public T DeserializeObject<T>(string inputString)
        {
            return _microsoftGraphSerializer.DeserializeObject<T>(inputString);
        }

        public string SerializeObject(object serializeableObject)
        {
            // If a Schema object is being serialized, do the conversion
            // ourselves
            if (serializeableObject is Schema)
            {
                var foo = JsonConvert.SerializeObject(serializeableObject, _jsonSerializerSettings);
                return foo;
            }

            // Otherwise, just pass through to the default Microsoft Graph SDK serializer
            return _microsoftGraphSerializer.SerializeObject(serializeableObject);
        }
    }
}