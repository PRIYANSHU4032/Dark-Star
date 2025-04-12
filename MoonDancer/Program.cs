using MoonDancer.Extracters;

var builder = WebApplication.CreateBuilder(args);

// Retrieve Maintainer path from appsettings.json
var maintainerPath = builder.Configuration["AppSettings:Maintainer"];
var excelPath = builder.Configuration["AppSettings:Excel_path"];

// Add services to the container
builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();
builder.Services.AddScoped<ProcessSyncManager>();
builder.Services.AddScoped<ExcelTableExtractor>();
builder.Services.AddScoped<MasterModuleManager>();
builder.Services.AddScoped<ReferenceIDsManager>();


var app = builder.Build();

// Configure the HTTP request pipeline
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();
app.UseAuthorization();
app.MapControllers();

app.Run();
